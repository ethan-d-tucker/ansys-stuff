"""
Ansys MAPDL MCP Server
======================
Gives Claude interactive, step-by-step control of Ansys MAPDL via MCP tools.
Uses FastMCP with stdio transport.
"""

import json
import logging
import os
import sys
from contextlib import asynccontextmanager
from typing import Any, Optional

import numpy as np
from fastmcp import FastMCP

# ---------------------------------------------------------------------------
# Logging — must go to stderr, stdout is reserved for MCP protocol
# ---------------------------------------------------------------------------
logging.basicConfig(
    stream=sys.stderr,
    level=logging.INFO,
    format="[ANSYS-MCP] %(levelname)s: %(message)s",
)
logger = logging.getLogger(__name__)


def _r(data: dict) -> str:
    """Serialize result dict to JSON string for MCP transport."""
    return json.dumps(data, indent=2, default=str)


# ---------------------------------------------------------------------------
# Global MAPDL session
# ---------------------------------------------------------------------------
_mapdl = None


def _get_mapdl():
    """Return the active MAPDL session or raise."""
    if _mapdl is None:
        raise RuntimeError("MAPDL is not running. Call launch_mapdl first.")
    return _mapdl


# ---------------------------------------------------------------------------
# Server setup with lifespan
# ---------------------------------------------------------------------------
@asynccontextmanager
async def lifespan(server):
    logger.info("Ansys MCP server starting")
    yield
    global _mapdl
    if _mapdl is not None:
        try:
            _mapdl.exit()
            logger.info("MAPDL session closed on shutdown")
        except Exception:
            pass
        _mapdl = None
    logger.info("Ansys MCP server stopped")


mcp = FastMCP(name="Ansys MAPDL Tools", lifespan=lifespan)


# ===========================  SESSION MANAGEMENT  ===========================

@mcp.tool()
def launch_mapdl(
    exec_file: Optional[str] = None,
    additional_switches: str = "-smp",
) -> str:
    """Launch an Ansys MAPDL session. Call this before any other tool.

    Args:
        exec_file: Optional full path to ANSYS executable. Auto-detected if omitted.
        additional_switches: MAPDL launch switches (default: -smp for shared memory parallel).
    """
    global _mapdl
    if _mapdl is not None:
        return _r({"status": "already_running", "version": str(_mapdl.version), "directory": _mapdl.directory})

    from ansys.mapdl.core import launch_mapdl as _launch

    kwargs = {"override": True, "loglevel": "WARNING", "additional_switches": additional_switches}
    if exec_file:
        kwargs["exec_file"] = exec_file

    logger.info("Launching MAPDL ...")
    _mapdl = _launch(**kwargs)
    logger.info(f"MAPDL {_mapdl.version} ready at {_mapdl.directory}")
    return _r({"status": "launched", "version": str(_mapdl.version), "directory": _mapdl.directory})


@mcp.tool()
def close_mapdl() -> str:
    """Save the database and close the MAPDL session."""
    global _mapdl
    m = _get_mapdl()
    try:
        m.finish()
        m.save()
    except Exception:
        pass
    m.exit()
    _mapdl = None
    return _r({"status": "closed"})


@mcp.tool()
def get_session_info() -> str:
    """Check whether MAPDL is running and return session info."""
    if _mapdl is None:
        return _r({"running": False})
    return _r({"running": True, "version": str(_mapdl.version), "directory": _mapdl.directory})


# ===========================  GEOMETRY  =====================================

@mcp.tool()
def import_geometry(file_path: str) -> str:
    """Import a CAD file (Parasolid .x_t, STEP .stp, IGES .igs) into MAPDL.

    Args:
        file_path: Full path to the geometry file.
    """
    m = _get_mapdl()
    m.prep7()

    file_path = os.path.normpath(file_path)
    if not os.path.isfile(file_path):
        return _r({"error": f"File not found: {file_path}"})

    basename = os.path.splitext(os.path.basename(file_path))[0]
    ext = os.path.splitext(file_path)[1].lstrip(".").lower()

    # Upload the file to the MAPDL working directory
    m.upload(file_path)
    work_dir = m.directory

    import_error = None

    # For Parasolid and STEP: use the external ac4 converter to create .anf,
    # then read it with mapdl.input() (Python-side, handles multi-line commands).
    # This works around gRPC limitations with ~PARAIN/~SATIN.
    if ext in ("x_t", "x_b", "xmt_txt", "stp", "step"):
        import subprocess

        # Pick the right converter
        ansys_base = r"C:\Program Files\ANSYS Inc\ANSYS Student\v252"
        if ext in ("x_t", "x_b", "xmt_txt"):
            converter = os.path.join(ansys_base, "ansys", "ac4", "bin", "para", "winx64", "ac4para.exe")
        else:
            converter = os.path.join(ansys_base, "commonfiles", "CAD", "ACIS", "winx64", "ac4sat.exe")

        if not os.path.isfile(converter):
            # Fallback: try the tilde command anyway
            try:
                if ext in ("x_t", "x_b", "xmt_txt"):
                    m.run(f"~PARAIN,'{basename}','{ext}',,,,0,0,1")
                else:
                    m.run(f"~SATIN,'{basename}','{ext}',,,,0")
            except Exception as e:
                import_error = str(e)
        else:
            # Run the converter externally — use relative filenames with cwd
            anf_name = f"{basename}.anf"
            src_name = f"{basename}.{ext}"
            anf_file = os.path.join(work_dir, anf_name)

            # Set up environment with Ansys paths for DLL resolution
            env = os.environ.copy()
            ansys_bin = os.path.join(ansys_base, "ansys", "bin", "winx64")
            para_bin = os.path.dirname(converter)
            env["PATH"] = f"{ansys_bin};{para_bin};{env.get('PATH', '')}"

            try:
                result = subprocess.run(
                    [converter, src_name, anf_name, "SOLIDS", "ANF"],
                    cwd=work_dir, env=env, timeout=60,
                    stdout=subprocess.PIPE, stderr=subprocess.PIPE,
                )
                logger.info(f"Converter return code: {result.returncode}")
                logger.info(f"Converter stdout: {result.stdout.decode(errors='replace')}")
                if result.stderr:
                    logger.warning(f"Converter stderr: {result.stderr.decode(errors='replace')}")
            except Exception as e:
                import_error = f"Converter failed: {e}"
                logger.error(f"Converter exception: {e}")

            # Read the ANF file using PyMAPDL's input() method
            if os.path.isfile(anf_file):
                try:
                    m.input(anf_file)
                    # ANF files use /aux15 internally; make sure we're back in PREP7
                    try:
                        m.finish()
                    except Exception:
                        pass
                    m.prep7()
                except Exception as e:
                    import_error = f"input() failed: {e}"
                    logger.warning(f"mapdl.input() raised: {e}")
            else:
                import_error = f"Converter did not produce {anf_file}"

    elif ext in ("igs", "iges"):
        try:
            m.run("/AUX15")
            m.run("IOPTN,MERGE,YES")
            m.run(f"IGESIN,'{basename}','{ext}'")
            m.run("FINISH")
            m.prep7()
        except Exception as e:
            import_error = str(e)
    else:
        return _r({"error": f"Unsupported file type: .{ext}"})

    if import_error:
        logger.warning(f"Import issue: {import_error}")

    m.allsel()
    n_vol = m.geometry.n_volu
    n_area = m.geometry.n_area
    n_kp = m.geometry.n_keypoint

    if n_vol == 0 and n_area == 0:
        return _r({"error": "No geometry imported. Check the file.", "details": import_error or ""})

    return _r({"status": "imported", "file": file_path, "volumes": n_vol, "areas": n_area, "keypoints": n_kp})


@mcp.tool()
def get_geometry_info() -> str:
    """Return geometry stats and bounding box of the current model."""
    m = _get_mapdl()
    m.allsel()

    info = {"volumes": m.geometry.n_volu, "areas": m.geometry.n_area, "keypoints": m.geometry.n_keypoint}

    try:
        if m.mesh.n_node > 0:
            coords = m.mesh.nodes
            mins = coords.min(axis=0)
            maxs = coords.max(axis=0)
            spans = maxs - mins
            info["bounding_box"] = {
                "x_min": float(mins[0]), "x_max": float(maxs[0]), "x_span": float(spans[0]),
                "y_min": float(mins[1]), "y_max": float(maxs[1]), "y_span": float(spans[1]),
                "z_min": float(mins[2]), "z_max": float(maxs[2]), "z_span": float(spans[2]),
            }
            info["nodes"] = m.mesh.n_node
            info["elements"] = m.mesh.n_elem
    except Exception:
        pass

    return _r(info)


# ===========================  MATERIALS  ====================================

@mcp.tool()
def define_material(mat_id: int, name: str, properties: dict) -> str:
    """Define a material with given properties.

    Args:
        mat_id: Material ID number (1, 2, 3, ...).
        name: Descriptive name for reference.
        properties: Dict of property labels and values. Supported keys:
            EX, EY, EZ (elastic moduli in Pa),
            GXY, GXZ, GYZ (shear moduli in Pa),
            PRXY, PRXZ, PRYZ (Poisson's ratios),
            DENS (density in kg/m3),
            ALPX, ALPY, ALPZ (CTE in 1/K).
    """
    m = _get_mapdl()
    m.prep7()

    valid_keys = {"EX", "EY", "EZ", "GXY", "GXZ", "GYZ",
                  "PRXY", "PRXZ", "PRYZ", "DENS", "ALPX", "ALPY", "ALPZ"}

    set_props = {}
    for key, val in properties.items():
        ku = key.upper()
        if ku in valid_keys:
            m.mp(ku, mat_id, val)
            set_props[ku] = val

    return _r({"status": "defined", "mat_id": mat_id, "name": name, "properties": set_props})


@mcp.tool()
def list_materials() -> str:
    """List all defined materials and their properties."""
    m = _get_mapdl()
    m.prep7()
    result = m.run("MPLIST,ALL")
    return _r({"material_list": result})


# ===========================  COMPOSITE SECTION  ============================

@mcp.tool()
def define_composite_section(section_id: int, name: str, plies: list[dict]) -> str:
    """Define a layered solid composite section.

    Args:
        section_id: Section ID number.
        name: Descriptive name.
        plies: List of ply dicts, each with keys:
            mat_id (int): Material ID for this ply.
            thickness (float): Ply thickness in metres.
            angle (float): Ply orientation angle in degrees.
            integration_points (int, optional): Number of integration points (default 3).
    """
    m = _get_mapdl()
    m.prep7()

    m.sectype(section_id, "SHELL", "", name)

    total_thickness = 0.0
    ply_summary = []
    for i, ply in enumerate(plies):
        mat = ply["mat_id"]
        t = ply["thickness"]
        angle = ply.get("angle", 0.0)
        nip = ply.get("integration_points", 3)
        m.secdata(t, mat, angle, nip)
        total_thickness += t
        ply_summary.append({"ply": i + 1, "mat_id": mat, "thickness_mm": round(t * 1000, 4), "angle_deg": angle})

    return _r({
        "status": "defined", "section_id": section_id, "name": name,
        "num_plies": len(plies), "total_thickness_mm": round(total_thickness * 1000, 4),
        "plies": ply_summary,
    })


# ===========================  ELEMENTS & MESHING  ===========================

@mcp.tool()
def set_element_type(type_id: int, element_name: str, keyopts: Optional[dict] = None) -> str:
    """Define an element type.

    Args:
        type_id: Element type ID (1, 2, ...).
        element_name: MAPDL element name (e.g., "SOLID186", "SHELL281").
        keyopts: Optional dict of KEYOPT number -> value (e.g., {3: 1} for layered solid).
    """
    m = _get_mapdl()
    m.prep7()

    m.et(type_id, element_name)
    set_keyopts = {}
    if keyopts:
        for k, v in keyopts.items():
            m.keyopt(type_id, int(k), v)
            set_keyopts[int(k)] = v

    return _r({"status": "defined", "type_id": type_id, "element": element_name, "keyopts": set_keyopts})


@mcp.tool()
def mesh_model(
    element_size: float = 0,
    smart_size: int = 4,
    shape: str = "tet",
    section_id: int = 1,
    type_id: int = 1,
    mat_id: int = 1,
) -> str:
    """Mesh all volumes in the model.

    Args:
        element_size: Target element size in metres. 0 = auto (SmartSize).
        smart_size: SmartSize level 1-10 (1=finest, 10=coarsest). Used when element_size=0.
        shape: Element shape — "tet" or "hex".
        section_id: Section ID to assign.
        type_id: Element type ID to assign.
        mat_id: Default material ID to assign.
    """
    m = _get_mapdl()
    m.prep7()
    m.allsel()

    m.mat(mat_id)
    m.type(type_id)
    m.secnum(section_id)

    if element_size > 0:
        m.esize(element_size)
    else:
        m.smrtsize(smart_size)

    if shape == "tet":
        m.mshape(1, "3D")
    else:
        m.mshape(0, "3D")
    m.mshkey(0)

    m.vmesh("ALL")

    return _r({"status": "meshed", "nodes": m.mesh.n_node, "elements": m.mesh.n_elem, "shape": shape})


@mcp.tool()
def get_mesh_info() -> str:
    """Return mesh statistics."""
    m = _get_mapdl()
    m.allsel()
    return _r({"nodes": m.mesh.n_node, "elements": m.mesh.n_elem})


# ===========================  SELECTION & BCs  ==============================

@mcp.tool()
def select_nodes(axis: str, min_val: float, max_val: float) -> str:
    """Select nodes within a coordinate range.

    Args:
        axis: Coordinate axis — "X", "Y", or "Z".
        min_val: Minimum coordinate value (metres).
        max_val: Maximum coordinate value (metres).
    """
    m = _get_mapdl()
    axis = axis.upper()
    if axis not in ("X", "Y", "Z"):
        return _r({"error": f"Invalid axis: {axis}. Must be X, Y, or Z."})

    m.nsel("S", "LOC", axis, min_val, max_val)
    return _r({"status": "selected", "axis": axis, "range": [min_val, max_val], "nodes_selected": m.mesh.n_node})


@mcp.tool()
def select_all() -> str:
    """Re-select all nodes, elements, and geometry entities."""
    m = _get_mapdl()
    m.allsel()
    return _r({"status": "all_selected", "nodes": m.mesh.n_node, "elements": m.mesh.n_elem})


@mcp.tool()
def apply_fixed_support() -> str:
    """Apply fixed support (all DOFs = 0) to the currently selected nodes."""
    m = _get_mapdl()
    n = m.mesh.n_node
    m.d("ALL", "ALL", 0)
    return _r({"status": "fixed_support_applied", "nodes_constrained": n})


@mcp.tool()
def apply_displacement(dof: str, value: float) -> str:
    """Apply a displacement constraint on the currently selected nodes.

    Args:
        dof: Degree of freedom — "UX", "UY", "UZ", "ROTX", "ROTY", "ROTZ", or "ALL".
        value: Displacement value.
    """
    m = _get_mapdl()
    m.d("ALL", dof.upper(), value)
    return _r({"status": "displacement_applied", "dof": dof.upper(), "value": value})


# ===========================  SOLVERS  ======================================

@mcp.tool()
def solve_modal(
    num_modes: int = 20,
    freq_start: float = 0.0,
    freq_end: float = 3000.0,
    solver: str = "LANB",
) -> str:
    """Run a modal analysis.

    Args:
        num_modes: Number of modes to extract.
        freq_start: Lower frequency bound (Hz).
        freq_end: Upper frequency bound (Hz).
        solver: Modal solver — "LANB" (Block Lanczos) or "UNSYM".
    """
    m = _get_mapdl()
    m.allsel()

    m.run("/SOLU")
    m.antype("MODAL")
    m.modopt(solver, num_modes, freq_start, freq_end)
    m.eqslv("SPARSE")
    m.mxpand(num_modes, 0, 0, "YES")

    logger.info(f"Solving modal: {num_modes} modes, {freq_start}-{freq_end} Hz")
    m.solve()
    m.finish()

    # Extract natural frequencies
    m.post1()
    frequencies = []
    for i in range(1, num_modes + 1):
        try:
            m.set(1, i)
            freq = m.get("FREQ_VAL", "ACTIVE", 0, "SET", "FREQ")
            frequencies.append({"mode": i, "frequency_hz": round(float(freq), 4)})
        except Exception:
            break
    m.finish()

    return _r({"status": "solved", "modes_found": len(frequencies), "frequencies": frequencies})


@mcp.tool()
def solve_psd(
    psd_table: list[dict],
    direction: str = "UY",
    damping_ratio: float = 0.02,
    num_modes: int = 20,
    select_axis: Optional[str] = None,
    select_min: Optional[float] = None,
    select_max: Optional[float] = None,
) -> str:
    """Run a PSD random vibration analysis (requires prior modal solve).

    Args:
        psd_table: List of dicts with "frequency_hz" and "psd_g2_per_hz" keys.
        direction: Excitation DOF — "UX", "UY", or "UZ".
        damping_ratio: Constant modal damping ratio (e.g., 0.02 = 2%).
        num_modes: Number of modes to include.
        select_axis: Axis for re-selecting BC nodes ("X", "Y", or "Z").
            If not provided, uses currently selected nodes.
        select_min: Min coordinate for node selection (metres).
        select_max: Max coordinate for node selection (metres).
    """
    m = _get_mapdl()
    direction = direction.upper()

    m.run("/SOLU")
    m.antype("SPECTR")
    m.spopt("PSD", num_modes, "ON")

    m.psdunit(1, "ACCG", 9.80665)

    freqs = [p["frequency_hz"] for p in psd_table]
    vals = [p["psd_g2_per_hz"] for p in psd_table]

    for i in range(0, len(freqs), 6):
        m.psdfrq(1, 1, *freqs[i:i + 6])
        m.psdval(1, *vals[i:i + 6])

    if select_axis and select_min is not None and select_max is not None:
        m.nsel("S", "LOC", select_axis.upper(), select_min, select_max)

    m.d("ALL", direction, 1.0)
    m.allsel()

    m.pfact(1, "BASE")
    m.psdres("DISP", 1)
    m.psdres("VELO", 1)
    m.psdres("ACEL", 1)
    m.psdres("STRE", 1)
    m.dmprat(damping_ratio)

    logger.info(f"Solving PSD: {direction}, damping={damping_ratio*100:.1f}%")
    m.solve()
    m.finish()

    return _r({
        "status": "solved", "direction": direction, "damping_ratio": damping_ratio,
        "psd_points": len(psd_table), "freq_range": f"{freqs[0]}-{freqs[-1]} Hz",
    })


# ===========================  POST-PROCESSING  =============================

@mcp.tool()
def get_natural_frequencies() -> str:
    """Extract natural frequencies from a completed modal analysis."""
    m = _get_mapdl()
    m.post1()

    frequencies = []
    for i in range(1, 100):
        try:
            m.set(1, i)
            freq = m.get("FREQ_VAL", "ACTIVE", 0, "SET", "FREQ")
            if float(freq) == 0.0 and i > 1:
                break
            frequencies.append({"mode": i, "frequency_hz": round(float(freq), 4)})
        except Exception:
            break

    m.finish()
    return _r({"modes": len(frequencies), "frequencies": frequencies})


@mcp.tool()
def get_displacement_results(component: str = "Y") -> str:
    """Get displacement results from the last solved analysis.

    Args:
        component: Displacement component — "X", "Y", "Z", or "NORM" (magnitude).
    """
    m = _get_mapdl()
    m.post1()

    try:
        m.set("LAST")
    except Exception:
        try:
            m.set(2, 1)
        except Exception:
            m.set(1, 1)

    m.allsel()
    component = component.upper()

    try:
        disp = m.post_processing.nodal_displacement("NORM" if component == "NORM" else component)
        max_val = float(np.max(np.abs(disp)))
        result = {
            "component": component,
            "max_absolute": max_val,
            "max_absolute_mm": round(max_val * 1000, 6),
            "max_absolute_um": round(max_val * 1e6, 4),
            "max_at_node_index": int(np.argmax(np.abs(disp))),
            "min": float(np.min(disp)),
            "max": float(np.max(disp)),
            "mean": float(np.mean(disp)),
        }
    except Exception as e:
        result_text = m.run(f"PRNSOL,U,{component}")
        result = {"component": component, "apdl_output": result_text}

    m.finish()
    return _r(result)


@mcp.tool()
def get_stress_results(component: str = "EQV") -> str:
    """Get stress results from the last solved analysis.

    Args:
        component: Stress component — "EQV" (von Mises), "X", "Y", "Z", "XY", "XZ", "YZ".
    """
    m = _get_mapdl()
    m.post1()

    try:
        m.set("LAST")
    except Exception:
        try:
            m.set(2, 1)
        except Exception:
            m.set(1, 1)

    m.allsel()
    component = component.upper()

    try:
        stress = m.post_processing.nodal_eqv_stress() if component == "EQV" else m.post_processing.nodal_component_stress(component)
        max_val = float(np.max(np.abs(stress)))
        result = {
            "component": component,
            "max_absolute_pa": max_val,
            "max_absolute_mpa": round(max_val / 1e6, 4),
            "max_at_node_index": int(np.argmax(np.abs(stress))),
            "min_pa": float(np.min(stress)),
            "max_pa": float(np.max(stress)),
        }
    except Exception as e:
        result_text = m.run(f"PRNSOL,S,{component}")
        result = {"component": component, "apdl_output": result_text}

    m.finish()
    return _r(result)


# ===========================  PLOTTING  ======================================

@mcp.tool()
def plot_results(
    result_type: str = "stress",
    component: str = "EQV",
    screenshot_path: Optional[str] = None,
) -> str:
    """Plot contour results on the mesh and save a screenshot.

    Uses PyMAPDL's built-in plotting which renders the actual 3D FEM mesh.

    Args:
        result_type: What to plot — "stress" or "displacement".
        component: For stress: "EQV" (von Mises), "X", "Y", "Z", "XY", "XZ", "YZ".
                   For displacement: "X", "Y", "Z", or "NORM" (magnitude).
        screenshot_path: Full path for the PNG screenshot. Auto-generated if omitted.
    """
    m = _get_mapdl()
    m.post1()

    try:
        m.set("LAST")
    except Exception:
        try:
            m.set(2, 1)
        except Exception:
            m.set(1, 1)

    m.allsel()
    result_type = result_type.lower()
    component = component.upper()

    if screenshot_path is None:
        screenshot_path = os.path.join(m.directory, f"plot_{result_type}_{component}.png")

    try:
        if result_type == "displacement":
            comp_arg = "NORM" if component == "NORM" else component
            disp = m.post_processing.nodal_displacement(comp_arg)
            m.post_processing.plot_nodal_displacement(
                comp_arg,
                title=f"Displacement ({component})",
                show_node_numbering=False,
                cpos="iso",
                screenshot=screenshot_path,
                off_screen=True,
            )
            max_val = float(np.max(np.abs(disp)))
            min_val = float(np.min(disp))
            max_display = f"{max_val * 1000:.6f} mm"
        elif result_type == "stress":
            if component == "EQV":
                stress = m.post_processing.nodal_eqv_stress()
                m.post_processing.plot_nodal_eqv_stress(
                    title="von Mises Stress",
                    show_node_numbering=False,
                    cpos="iso",
                    screenshot=screenshot_path,
                    off_screen=True,
                )
            else:
                stress = m.post_processing.nodal_component_stress(component)
                m.post_processing.plot_nodal_component_stress(
                    component,
                    title=f"Stress ({component})",
                    show_node_numbering=False,
                    cpos="iso",
                    screenshot=screenshot_path,
                    off_screen=True,
                )
            max_val = float(np.max(np.abs(stress)))
            min_val = float(np.min(stress))
            max_display = f"{max_val / 1e6:.4f} MPa"
        else:
            m.finish()
            return _r({"error": f"Unknown result_type '{result_type}'. Use 'stress' or 'displacement'."})

        m.finish()
        return _r({
            "screenshot": screenshot_path,
            "result_type": result_type,
            "component": component,
            "max_absolute": max_val,
            "max_display": max_display,
            "min": min_val,
            "max": float(np.max(stress if result_type == "stress" else disp)),
        })

    except Exception as e:
        # Fallback: use APDL native plotting (won't save screenshot but will display in MAPDL GUI)
        logger.warning(f"PyVista plot failed ({e}), falling back to APDL /PLNSOL")
        try:
            if result_type == "displacement":
                apdl_comp = "SUM" if component == "NORM" else component
                m.run(f"PLNSOL,U,{apdl_comp},0,1")
            else:
                m.run(f"PLNSOL,S,{component},0,1")
            m.finish()
            return _r({
                "fallback": "apdl_plnsol",
                "note": "PyVista unavailable; used MAPDL native plotter (no screenshot saved)",
                "error_detail": str(e),
            })
        except Exception as e2:
            m.finish()
            return _r({"error": f"Both PyVista and APDL plotting failed: {e}; {e2}"})


# ===========================  ESCAPE HATCH  =================================

@mcp.tool()
def run_apdl_command(command: str) -> str:
    """Run an arbitrary APDL command and return its output.

    Args:
        command: The APDL command string to execute (e.g., "NLIST" or "/PREP7").
    """
    m = _get_mapdl()
    try:
        output = m.run(command)
        return _r({"command": command, "output": output})
    except Exception as e:
        return _r({"command": command, "error": str(e)})


@mcp.tool()
def run_apdl_commands(commands: list[str]) -> str:
    """Run multiple APDL commands sequentially.

    Args:
        commands: List of APDL command strings to execute in order.
    """
    m = _get_mapdl()
    results = []
    for cmd in commands:
        try:
            output = m.run(cmd)
            results.append({"command": cmd, "output": output, "status": "ok"})
        except Exception as e:
            results.append({"command": cmd, "error": str(e), "status": "error"})
    return _r({"results": results})


# ===========================  ENTRY POINT  ==================================

if __name__ == "__main__":
    mcp.run()
