"""
Parameterised PSD simulation engine.

Refactored from run_simulation.py to accept arbitrary geometry, materials,
PSD profiles, and excitation directions.  The modal solve runs once; the
PSD SRSS computation can be repeated cheaply for different profiles / axes.

All proven patterns preserved: ac4 Parasolid import, SOLID187 free tet mesh,
manual SRSS 1-sigma computation (Student edition workaround).
"""

import os
import sys
import time
import shutil
import subprocess
from dataclasses import dataclass, field

import numpy as np
from ansys.mapdl.core import launch_mapdl

from composite_failure import signed_srss


# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------

ANSYS_BASE = r"C:\Program Files\ANSYS Inc\ANSYS Student\v252"


@dataclass
class SimulationConfig:
    """All parameters needed to run one PSD analysis."""
    geometry_file: str
    part_name: str = "Part"
    element_size: float = 0.003           # metres
    material_props: dict = field(default_factory=dict)
    material_name: str = "Carbon/Epoxy"
    psd_table: list = field(default_factory=list)   # [(freq, g2/Hz), ...]
    excitation_direction: str = "Y"
    damping_ratio: float = 0.02
    num_modes: int = 20
    freq_start: float = 0.0
    freq_end: float = 3000.0


@dataclass
class ModalData:
    """Intermediate data from the modal solve (reusable across PSD runs)."""
    nat_freqs: np.ndarray               # (n_modes,)
    n_modes: int
    mode_disp: dict                     # {"X": (n_modes, n_nodes), "Y": ..., "Z": ...}
    mode_stress: dict                   # {"SX": ..., "SY": ..., "SZ": ..., "SXY": ..., "SXZ": ..., "SYZ": ..., "EQV": ...}
    coords: np.ndarray                  # (n_nodes, 3)
    all_nids: np.ndarray                # (n_nodes,)
    free_mask: np.ndarray               # (n_nodes,) bool
    mesh_info: dict
    bc_info: dict
    solve_time: float
    participation: dict = field(default_factory=dict)  # {"X": (n_modes,), "Y": ..., "Z": ...}
    grid_base: object = None            # PyVista grid if available


@dataclass
class PSDResults:
    """Results from one PSD computation (one profile, one axis)."""
    config_label: str                   # e.g. "MIN_INTEGRITY / Y-axis"
    psd_table: list
    excitation_direction: str
    damping_ratio: float

    modal_sigma2: np.ndarray            # (n_modes,)  modal variance

    # 1-sigma displacement (unsigned magnitude via SRSS)
    disp_x: np.ndarray
    disp_y: np.ndarray
    disp_z: np.ndarray
    disp_mag: np.ndarray

    # 1-sigma stress components (signed via dominant-mode convention)
    stress_sx: np.ndarray
    stress_sy: np.ndarray
    stress_sz: np.ndarray
    stress_sxy: np.ndarray
    stress_sxz: np.ndarray
    stress_syz: np.ndarray
    stress_eqv: np.ndarray              # unsigned von Mises

    input_grms: float

    # Peak values (over free nodes only)
    max_disp_mag_um: float
    max_stress_eqv_mpa: float


# ---------------------------------------------------------------------------
# Timing helper
# ---------------------------------------------------------------------------

_t0 = None

def _elapsed():
    return f"{time.time() - _t0:.1f}s" if _t0 else ""


# ---------------------------------------------------------------------------
# Step 1 + 2: Launch MAPDL and import geometry
# ---------------------------------------------------------------------------

def launch_and_import(config, mapdl=None):
    """
    Launch MAPDL (if not provided) and import geometry via ac4 converter.

    Returns (mapdl, geometry_info_dict).
    """
    global _t0
    _t0 = time.time()

    if mapdl is None:
        print(f"  Launching MAPDL ...")
        mapdl = launch_mapdl(override=True, loglevel="WARNING", start_timeout=120)
        print(f"  MAPDL v{mapdl.version}")

    work_dir = mapdl.directory
    geo_file = os.path.abspath(config.geometry_file)
    ext = os.path.splitext(geo_file)[1].lower()
    basename = os.path.basename(geo_file)

    print(f"  Importing geometry: {basename}")

    # Copy geometry to working directory
    dst = os.path.join(work_dir, basename)
    shutil.copy2(geo_file, dst)

    # Parasolid via ac4 converter
    if ext in (".x_t", ".x_b", ".xmt_txt"):
        converter = os.path.join(ANSYS_BASE, "ansys", "ac4", "bin", "para",
                                 "winx64", "ac4para.exe")
        schema_dir = os.path.join(ANSYS_BASE, "commonfiles", "CAD", "Siemens",
                                  "Parasolid36.1.227", "winx64", "schema")
        env = os.environ.copy()
        env["PATH"] = (
            os.path.join(ANSYS_BASE, "ansys", "bin", "winx64") + ";" +
            os.path.dirname(converter) + ";" +
            os.path.join(ANSYS_BASE, "commonfiles", "CAD", "Siemens",
                         "Parasolid36.1.227", "winx64") + ";" +
            env.get("PATH", "")
        )
        env["P_SCHEMA"] = schema_dir

        anf_name = os.path.splitext(basename)[0] + ".anf"
        result = subprocess.run(
            [converter, basename, anf_name, "SOLIDS", "ANF"],
            cwd=work_dir, env=env, timeout=120,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        if result.returncode != 0:
            raise RuntimeError(
                f"ac4 converter failed (rc={result.returncode}): "
                f"{result.stdout.decode(errors='replace')}"
            )
        anf_file = os.path.join(work_dir, anf_name)
        print(f"  ac4 converter: {os.path.getsize(anf_file)} bytes")

    # STEP files via ac4sat
    elif ext in (".stp", ".step"):
        converter = os.path.join(ANSYS_BASE, "ansys", "ac4", "bin", "sat",
                                 "winx64", "ac4sat.exe")
        env = os.environ.copy()
        env["PATH"] = (
            os.path.join(ANSYS_BASE, "ansys", "bin", "winx64") + ";" +
            os.path.dirname(converter) + ";" +
            env.get("PATH", "")
        )
        anf_name = os.path.splitext(basename)[0] + ".anf"
        result = subprocess.run(
            [converter, basename, anf_name, "SOLIDS", "ANF"],
            cwd=work_dir, env=env, timeout=120,
            stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        )
        if result.returncode != 0:
            raise RuntimeError(f"ac4sat converter failed: {result.stdout.decode(errors='replace')}")
        anf_file = os.path.join(work_dir, anf_name)
    else:
        raise ValueError(f"Unsupported geometry format: {ext}")

    # Load into MAPDL
    mapdl.prep7()
    mapdl.units("SI")
    mapdl.input(anf_file)
    try:
        mapdl.finish()
    except Exception:
        pass
    mapdl.prep7()
    mapdl.allsel()

    nv = mapdl.geometry.n_volu
    na = mapdl.geometry.n_area
    print(f"  Geometry loaded: {nv} vol, {na} areas")

    return mapdl, {"n_volu": nv, "n_area": na, "file": basename}


# ---------------------------------------------------------------------------
# Step 3-5: Material, Mesh, BCs
# ---------------------------------------------------------------------------

def setup_model(mapdl, config):
    """
    Define element type, material, mesh, and boundary conditions.

    Returns (mesh_info, bc_info, coords, all_nids, free_mask).
    """
    print(f"  Setting element type + material: {config.material_name}")
    mapdl.et(1, "SOLID187")
    for prop, val in config.material_props.items():
        mapdl.mp(prop, 1, val)

    # Mesh
    print(f"  Meshing (esize={config.element_size*1000:.1f} mm) ...")
    mapdl.allsel()
    mapdl.mat(1)
    mapdl.type(1)
    mapdl.esize(config.element_size)
    mapdl.mshape(1, "3D")
    mapdl.mshkey(0)
    mapdl.vmesh("ALL")

    nn = mapdl.mesh.n_node
    ne = mapdl.mesh.n_elem
    print(f"  Mesh: {nn} nodes, {ne} elements")

    coords = mapdl.mesh.nodes
    mins = coords.min(axis=0)
    maxs = coords.max(axis=0)
    spans = maxs - mins

    mesh_info = {
        "n_nodes": nn,
        "n_elements": ne,
        "mins_m": mins.tolist(),
        "maxs_m": maxs.tolist(),
        "spans_m": spans.tolist(),
    }

    # Boundary conditions -- fix handle end (5% of longest axis)
    longest_idx = int(np.argmax(spans))
    axis_names = ["X", "Y", "Z"]
    long_axis = axis_names[longest_idx]
    tol = 0.05 * spans[longest_idx]
    fix_min = mins[longest_idx]

    mapdl.nsel("S", "LOC", long_axis, fix_min, fix_min + tol)
    n_fixed = mapdl.mesh.n_node
    mapdl.d("ALL", "ALL", 0)
    mapdl.allsel()
    print(f"  Fixed {n_fixed} nodes at {long_axis} min end")

    bc_info = {
        "axis": long_axis,
        "n_fixed": n_fixed,
        "fix_min": float(fix_min),
        "fix_tol": float(tol),
    }

    # Free-node mask
    all_nids = mapdl.mesh.nnum
    mapdl.prep7()
    mapdl.nsel("S", "LOC", long_axis, fix_min, fix_min + tol)
    fixed_set = set(mapdl.mesh.nnum)
    mapdl.allsel()
    free_mask = np.array([n not in fixed_set for n in all_nids])

    return mesh_info, bc_info, coords, all_nids, free_mask


# ---------------------------------------------------------------------------
# Step 6-7: Modal solve + extract modal data
# ---------------------------------------------------------------------------

def run_modal_solve(mapdl, config, coords, all_nids, free_mask, mesh_info, bc_info):
    """
    Run modal analysis and extract mode shapes + all stress components.

    Returns ModalData.
    """
    print(f"  Modal analysis: {config.num_modes} modes, "
          f"{config.freq_start:.0f}-{config.freq_end:.0f} Hz ...")

    mapdl.run("/SOLU")
    mapdl.antype("MODAL")
    mapdl.modopt("LANB", config.num_modes, config.freq_start, config.freq_end)
    mapdl.eqslv("SPARSE")
    mapdl.mxpand(config.num_modes, 0, 0, "YES")

    t_start = time.time()
    mapdl.solve()
    mapdl.finish()
    solve_time = time.time() - t_start
    print(f"  Solve time: {solve_time:.1f}s")

    # ---- Extract per-mode data ----
    print("  Extracting modal data (displacements + 6 stress components) ...")
    mapdl.post1()

    nat_freqs = []
    disp_modes = {"X": [], "Y": [], "Z": []}
    stress_modes = {"SX": [], "SY": [], "SZ": [], "SXY": [], "SXZ": [], "SYZ": [], "EQV": []}

    for mode in range(1, config.num_modes + 1):
        try:
            mapdl.set(1, mode)
            freq = mapdl.get("FREQ_VAL", "ACTIVE", 0, "SET", "FREQ")
            if freq <= 0:
                break
            nat_freqs.append(freq)

            # Displacements
            for comp in ("X", "Y", "Z"):
                d = mapdl.post_processing.nodal_displacement(comp)
                disp_modes[comp].append(d.copy())

            # Von Mises stress
            try:
                seqv = mapdl.post_processing.nodal_eqv_stress()
                stress_modes["EQV"].append(seqv.copy())
            except Exception:
                stress_modes["EQV"].append(np.zeros(len(disp_modes["X"][-1])))

            # Component stresses (SX, SY, SZ, SXY, SXZ, SYZ)
            for scomp in ("X", "Y", "Z", "XY", "XZ", "YZ"):
                key = f"S{scomp}"
                try:
                    sc = mapdl.post_processing.nodal_component_stress(scomp)
                    stress_modes[key].append(sc.copy())
                except Exception:
                    stress_modes[key].append(np.zeros(len(disp_modes["X"][-1])))

            print(f"    Mode {mode:2d}: {freq:10.2f} Hz")
        except Exception:
            break

    mapdl.finish()

    n_modes = len(nat_freqs)
    nat_freqs_arr = np.array(nat_freqs)

    mode_disp = {k: np.array(v[:n_modes]) for k, v in disp_modes.items()}
    mode_stress = {k: np.array(v[:n_modes]) for k, v in stress_modes.items()}

    # Try to get PyVista grid for contour plotting
    grid_base = None
    try:
        grid_base = mapdl.mesh.grid.copy()
    except Exception:
        pass

    # Compute modal participation factors for each direction.
    # For base excitation in direction d, the modal force on mode i is
    # proportional to L_i,d = sum(m_j * phi_{j,i,d}).  With uniform lumped
    # mass the per-node mass cancels in the directional weight, so we use
    # the raw sum of mode-shape values.  The directional weight distributes
    # each mode's response across X/Y/Z proportionally to L_i,d^2.
    participation = {}
    for d in ("X", "Y", "Z"):
        participation[d] = np.array([
            np.sum(mode_disp[d][i, :]) for i in range(n_modes)
        ])

    print(f"  Extracted {n_modes} modes")

    return ModalData(
        nat_freqs=nat_freqs_arr,
        n_modes=n_modes,
        mode_disp=mode_disp,
        mode_stress=mode_stress,
        coords=coords,
        all_nids=all_nids,
        free_mask=free_mask,
        mesh_info=mesh_info,
        bc_info=bc_info,
        solve_time=solve_time,
        participation=participation,
        grid_base=grid_base,
    )


# ---------------------------------------------------------------------------
# Step 8: PSD SRSS computation (pure numpy, fast, repeatable)
# ---------------------------------------------------------------------------

def compute_psd(modal_data, psd_table, excitation_direction="Y",
                damping_ratio=0.02, label=""):
    """
    Compute 1-sigma PSD response from pre-extracted modal data.

    This is the same proven SRSS algorithm from run_simulation.py,
    extended to produce all 6 signed stress components.

    Parameters
    ----------
    modal_data : ModalData
    psd_table : list of (freq_hz, g2_per_hz)
    excitation_direction : str  "X", "Y", or "Z"
    damping_ratio : float
    label : str  human-readable label for print output

    Returns
    -------
    PSDResults
    """
    g = 9.80665
    nat_freqs = modal_data.nat_freqs
    n_modes = modal_data.n_modes
    free_mask = modal_data.free_mask

    # Interpolate PSD onto log-spaced grid
    psd_freqs = np.array([f for f, _ in psd_table])
    psd_vals = np.array([v for _, v in psd_table]) * g**2  # -> (m/s^2)^2/Hz

    f_min = max(psd_freqs[0], 1.0)
    f_max = psd_freqs[-1]
    f_grid = np.logspace(np.log10(f_min), np.log10(f_max), 2000)

    log_psd = np.interp(np.log10(f_grid), np.log10(psd_freqs), np.log10(psd_vals))
    psd_interp = 10.0 ** log_psd

    # Modal variance (base integral, direction-independent)
    modal_sigma2_base = np.zeros(n_modes)
    for i, fi in enumerate(nat_freqs):
        omega_i = 2.0 * np.pi * fi
        r = f_grid / fi
        H2 = 1.0 / ((1.0 - r**2)**2 + (2.0 * damping_ratio * r)**2)
        modal_sigma2_base[i] = np.trapezoid(H2 * psd_interp, f_grid) / omega_i**4

    # Apply directional participation factor weighting.
    # Weight_i = PF_i,d^2 / (PF_i,X^2 + PF_i,Y^2 + PF_i,Z^2)
    # This distributes each mode's response by how much it participates
    # in the excitation direction, preserving total magnitude for isotropic
    # excitation while differentiating X vs Y vs Z.
    modal_sigma2 = modal_sigma2_base.copy()
    if modal_data.participation:
        pf_d = modal_data.participation.get(excitation_direction)
        if pf_d is not None:
            pf_total_sq = (
                modal_data.participation["X"] ** 2
                + modal_data.participation["Y"] ** 2
                + modal_data.participation["Z"] ** 2
            )
            for i in range(n_modes):
                if pf_total_sq[i] > 0:
                    modal_sigma2[i] *= pf_d[i] ** 2 / pf_total_sq[i]
                else:
                    modal_sigma2[i] *= 1.0 / 3.0  # isotropic fallback

    # 1-sigma displacement (unsigned SRSS)
    disp = {}
    for comp in ("X", "Y", "Z"):
        shapes = modal_data.mode_disp[comp]  # (n_modes, n_nodes)
        sigma2 = np.zeros(shapes.shape[1])
        for i in range(n_modes):
            sigma2 += shapes[i, :] ** 2 * modal_sigma2[i]
        disp[comp] = np.sqrt(sigma2)

    disp_mag = np.sqrt(disp["X"]**2 + disp["Y"]**2 + disp["Z"]**2)

    # 1-sigma stress -- SIGNED via dominant-mode convention
    stress = {}
    for scomp in ("SX", "SY", "SZ", "SXY", "SXZ", "SYZ"):
        stress[scomp] = signed_srss(modal_data.mode_stress[scomp], modal_sigma2)

    # Von Mises (unsigned SRSS)
    eqv_shapes = modal_data.mode_stress["EQV"]
    sigma2_eqv = np.zeros(eqv_shapes.shape[1])
    for i in range(n_modes):
        sigma2_eqv += eqv_shapes[i, :] ** 2 * modal_sigma2[i]
    stress_eqv = np.sqrt(sigma2_eqv)

    # Input Grms
    psd_freqs_raw = np.array([f for f, _ in psd_table])
    psd_vals_raw = np.array([v for _, v in psd_table])
    fg = np.logspace(np.log10(max(psd_freqs_raw[0], 1.0)),
                     np.log10(psd_freqs_raw[-1]), 2000)
    lp = np.interp(np.log10(fg), np.log10(psd_freqs_raw), np.log10(psd_vals_raw))
    input_grms = float(np.sqrt(np.trapezoid(10.0 ** lp, fg)))

    # Peak values over free nodes
    max_disp = float(np.max(disp_mag[free_mask])) * 1e6 if np.any(free_mask) else 0
    max_seqv = float(np.max(stress_eqv[free_mask])) / 1e6 if np.any(free_mask) else 0

    if label:
        print(f"    {label}: max|U|={max_disp:.2f} um, "
              f"max SEQV={max_seqv:.4f} MPa, Grms={input_grms:.3f}")

    return PSDResults(
        config_label=label,
        psd_table=psd_table,
        excitation_direction=excitation_direction,
        damping_ratio=damping_ratio,
        modal_sigma2=modal_sigma2,
        disp_x=disp["X"],
        disp_y=disp["Y"],
        disp_z=disp["Z"],
        disp_mag=disp_mag,
        stress_sx=stress["SX"],
        stress_sy=stress["SY"],
        stress_sz=stress["SZ"],
        stress_sxy=stress["SXY"],
        stress_sxz=stress["SXZ"],
        stress_syz=stress["SYZ"],
        stress_eqv=stress_eqv,
        input_grms=input_grms,
        max_disp_mag_um=max_disp,
        max_stress_eqv_mpa=max_seqv,
    )


# ---------------------------------------------------------------------------
# Full end-to-end run
# ---------------------------------------------------------------------------

def run_full_analysis(config, mapdl=None):
    """
    Run the full PSD analysis for a single config.

    Returns (mapdl, ModalData, PSDResults).
    """
    mapdl, geo_info = launch_and_import(config, mapdl)
    mesh_info, bc_info, coords, all_nids, free_mask = setup_model(mapdl, config)
    mesh_info["geometry"] = geo_info
    modal_data = run_modal_solve(mapdl, config, coords, all_nids, free_mask,
                                 mesh_info, bc_info)
    psd_results = compute_psd(
        modal_data, config.psd_table,
        excitation_direction=config.excitation_direction,
        damping_ratio=config.damping_ratio,
        label=f"{config.part_name} / {config.excitation_direction}-axis",
    )
    return mapdl, modal_data, psd_results


def run_multi_environment(config, profiles, axes=None):
    """
    Run modal solve once, then compute PSD for each profile x axis combo.

    Parameters
    ----------
    config : SimulationConfig  (psd_table / excitation_direction ignored,
             overridden by each profile / axis)
    profiles : list[dict]  from mil_std_profiles.get_profile()
    axes : list[str] or None  default ["X", "Y", "Z"]

    Returns
    -------
    (mapdl, ModalData, results_dict)
        results_dict = {profile_id: {axis: PSDResults}}
    """
    if axes is None:
        axes = ["X", "Y", "Z"]

    # Step 1-7: model setup + modal solve (ONCE)
    mapdl, geo_info = launch_and_import(config)
    mesh_info, bc_info, coords, all_nids, free_mask = setup_model(mapdl, config)
    mesh_info["geometry"] = geo_info
    modal_data = run_modal_solve(mapdl, config, coords, all_nids, free_mask,
                                 mesh_info, bc_info)

    # Step 8: PSD for each profile x axis
    results = {}
    for prof in profiles:
        pid = prof["id"]
        results[pid] = {}
        print(f"\n  --- Environment: {prof['name']} ({prof['requirement_id']}) ---")
        for axis in axes:
            label = f"{pid} / {axis}-axis"
            psd_res = compute_psd(
                modal_data,
                psd_table=prof["psd_table"],
                excitation_direction=axis,
                damping_ratio=config.damping_ratio,
                label=label,
            )
            results[pid][axis] = psd_res

    return mapdl, modal_data, results
