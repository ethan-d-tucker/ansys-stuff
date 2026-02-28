"""
Composite Wrench Random Vibration (PSD) Analysis using PyMAPDL
==============================================================

This script performs a full random vibration analysis on a heavy-duty wrench
modeled as a composite sandwich structure with solid layered elements.

Analysis workflow:
  1. Import Parasolid geometry
  2. Define orthotropic composite materials (carbon prepreg + honeycomb core)
  3. Define layered solid section with symmetric sandwich layup
  4. Mesh with SOLID186 layered elements
  5. Apply fixed boundary conditions at the handle end
  6. Modal analysis (Block Lanczos, 20 modes, 0-3000 Hz)
  7. PSD spectrum analysis (base excitation in Y)
  8. Post-process 1-sigma displacement and stress results

Author: Generated for Ansys Student 2025 R2
"""

import sys
import time
import numpy as np

# ---------------------------------------------------------------------------
# Attempt to import PyMAPDL; exit gracefully if not installed.
# ---------------------------------------------------------------------------
try:
    from ansys.mapdl.core import launch_mapdl
except ImportError:
    print("ERROR: ansys-mapdl-core is not installed.")
    print("Install with:  pip install ansys-mapdl-core")
    sys.exit(1)

# Optional: import plotting utilities (PyVista) for post-processing figures.
try:
    import pyvista as pv
    HAS_PYVISTA = True
except ImportError:
    HAS_PYVISTA = False
    print("WARNING: pyvista not found. Plots will use MAPDL internal plotter.")


# ===========================  CONFIGURATION  ===============================

# Path to Parasolid geometry file
PARASOLID_FILE = r"C:\Users\EthanTucker\Downloads\heavyDutyWrench.x_t"

# Material 1 -- Epoxy carbon woven prepreg (fabric 0.286 mm ply)
MAT1 = {
    "EX":  60.0e9,     # Pa -- in-plane stiffness direction 1
    "EY":  60.0e9,     # Pa -- in-plane stiffness direction 2
    "EZ":  10.0e9,     # Pa -- through-thickness stiffness
    "GXY":  5.0e9,     # Pa -- in-plane shear
    "GXZ":  4.0e9,     # Pa -- transverse shear 13
    "GYZ":  4.0e9,     # Pa -- transverse shear 23
    "PRXY": 0.04,      # Poisson's ratio 12
    "PRXZ": 0.30,      # Poisson's ratio 13
    "PRYZ": 0.30,      # Poisson's ratio 23
    "DENS": 1420.0,    # kg/m^3
}

# Material 2 -- Honeycomb core (Nomex-style)
MAT2 = {
    "EX":   1.0e6,     # Pa -- very compliant in-plane
    "EY":   1.0e6,     # Pa
    "EZ": 130.0e6,     # Pa -- stiff through-thickness
    "GXY":  1.0e6,     # Pa
    "GXZ": 24.0e6,     # Pa -- transverse shear (ribbon direction)
    "GYZ": 48.0e6,     # Pa -- transverse shear (width direction)
    "PRXY": 0.49,
    "PRXZ": 0.001,
    "PRYZ": 0.001,
    "DENS": 48.0,      # kg/m^3
}

# Composite layup (symmetric sandwich)
# Each tuple: (material_id, thickness_mm, orientation_deg)
# Layup from bottom face to top face:
#   2 x carbon at 0 deg  |  2 x carbon at 45 deg  |  1 x carbon at 90 deg
#   honeycomb core
#   1 x carbon at 90 deg |  2 x carbon at 45 deg  |  2 x carbon at 0 deg
PLY_THICKNESS_CARBON = 0.286e-3   # metres (0.286 mm)
CORE_THICKNESS       = 6.35e-3    # metres (6.35 mm)

LAYUP = [
    # Bottom face-sheet (5 plies)
    (1, PLY_THICKNESS_CARBON,  0.0),   # ply  1 -- carbon 0 deg
    (1, PLY_THICKNESS_CARBON,  0.0),   # ply  2 -- carbon 0 deg
    (1, PLY_THICKNESS_CARBON, 45.0),   # ply  3 -- carbon 45 deg
    (1, PLY_THICKNESS_CARBON, 45.0),   # ply  4 -- carbon 45 deg
    (1, PLY_THICKNESS_CARBON, 90.0),   # ply  5 -- carbon 90 deg
    # Core
    (2, CORE_THICKNESS,        0.0),   # ply  6 -- honeycomb core
    # Top face-sheet (5 plies, symmetric)
    (1, PLY_THICKNESS_CARBON, 90.0),   # ply  7 -- carbon 90 deg
    (1, PLY_THICKNESS_CARBON, 45.0),   # ply  8 -- carbon 45 deg
    (1, PLY_THICKNESS_CARBON, 45.0),   # ply  9 -- carbon 45 deg
    (1, PLY_THICKNESS_CARBON,  0.0),   # ply 10 -- carbon 0 deg
    (1, PLY_THICKNESS_CARBON,  0.0),   # ply 11 -- carbon 0 deg
]

# Modal analysis settings
NUM_MODES    = 20
FREQ_START   = 0.0
FREQ_END     = 3000.0

# PSD base excitation table  (frequency [Hz], PSD value [G^2/Hz])
PSD_TABLE = [
    (  20.0, 0.010),
    (  80.0, 0.040),
    ( 350.0, 0.040),
    (2000.0, 0.007),
]

# Mesh element size -- adjust if needed for your geometry scale.
# A value of 0 lets MAPDL decide automatically based on geometry.
ELEMENT_SIZE = 0  # metres (0 = auto)


# ===========================================================================
#                       HELPER FUNCTIONS
# ===========================================================================

def print_banner(msg):
    """Print a highlighted status banner."""
    border = "=" * (len(msg) + 4)
    print(f"\n{border}")
    print(f"  {msg}")
    print(f"{border}\n")


def elapsed(t0):
    """Return formatted elapsed time since t0."""
    dt = time.time() - t0
    mins, secs = divmod(dt, 60)
    return f"{int(mins):02d}:{secs:05.2f}"


# ===========================================================================
#                       MAIN ANALYSIS
# ===========================================================================

def main():
    t_start = time.time()

    # ------------------------------------------------------------------
    # 1.  LAUNCH MAPDL
    # ------------------------------------------------------------------
    print_banner("Step 1: Launching MAPDL")
    print("Starting Ansys MAPDL session (Student 2025) ...")

    # launch_mapdl() will auto-detect the Ansys installation.
    # For Student 2025 the default port and mode should work.
    # Adjust 'exec_file' or 'additional_switches' if your install is
    # in a non-standard location.
    mapdl = launch_mapdl(
        run_location=None,          # use a temporary directory
        override=True,              # kill any zombie MAPDL processes
        loglevel="WARNING",         # reduce console noise
        additional_switches="-smp", # shared-memory parallel for speed
    )

    print(f"MAPDL version : {mapdl.version}")
    print(f"Working dir   : {mapdl.directory}")

    # ------------------------------------------------------------------
    # 2.  PREPROCESSOR -- IMPORT GEOMETRY
    # ------------------------------------------------------------------
    print_banner("Step 2: Importing Parasolid geometry")

    mapdl.prep7()
    mapdl.units("SI")  # Ensure SI unit system (metres, kg, seconds, Pa)

    # Import the Parasolid file.
    # ~PARAIN is the standard APDL command for Parasolid import.
    # Arguments: filename, extension, path, entity-type (0=auto),
    #            blank, blank, merge-coincident-keypoints (1=yes)
    # We split the path and filename for the command.
    import os
    para_dir  = os.path.dirname(PARASOLID_FILE).replace("\\", "/")
    para_name = os.path.splitext(os.path.basename(PARASOLID_FILE))[0]
    para_ext  = os.path.splitext(os.path.basename(PARASOLID_FILE))[1].lstrip(".")

    print(f"File: {PARASOLID_FILE}")
    print(f"  Directory : {para_dir}")
    print(f"  Name      : {para_name}")
    print(f"  Extension : {para_ext}")

    # Use ~PARAIN for Parasolid import
    mapdl.run(f"~PARAIN,'{para_name}','{para_ext}','{para_dir}/',,0,0,1")

    # Verify geometry was imported
    mapdl.allsel()
    num_vols   = mapdl.geometry.n_volu
    num_areas  = mapdl.geometry.n_area
    num_kps    = mapdl.geometry.n_keypoint
    print(f"Imported geometry: {num_vols} volume(s), {num_areas} area(s), "
          f"{num_kps} keypoint(s)")

    if num_vols == 0:
        print("ERROR: No volumes imported. Check the Parasolid file path.")
        mapdl.exit()
        sys.exit(1)

    print(f"  [{elapsed(t_start)}] Geometry import complete.")

    # ------------------------------------------------------------------
    # 3.  DEFINE ELEMENT TYPE
    # ------------------------------------------------------------------
    print_banner("Step 3: Defining element type -- SOLID186 layered")

    # SOLID186 is a 20-node higher-order 3-D solid element.
    # KEYOPT(3) = 1 activates the layered-solid formulation so that
    # composite section data (SECTYPE/SECDATA) is honoured.
    mapdl.et(1, "SOLID186")
    mapdl.keyopt(1, 3, 1)   # layered structural solid

    # KEYOPT(8) controls layer data storage:
    #   0 = store data for bottom & top of each layer (default)
    mapdl.keyopt(1, 8, 0)

    print("Element type 1: SOLID186  KEYOPT(3)=1 (layered solid)")

    # ------------------------------------------------------------------
    # 4.  DEFINE MATERIALS
    # ------------------------------------------------------------------
    print_banner("Step 4: Defining composite materials")

    # --- Material 1: Epoxy carbon woven prepreg ---
    mapdl.mp("EX",   1, MAT1["EX"])
    mapdl.mp("EY",   1, MAT1["EY"])
    mapdl.mp("EZ",   1, MAT1["EZ"])
    mapdl.mp("GXY",  1, MAT1["GXY"])
    mapdl.mp("GXZ",  1, MAT1["GXZ"])
    mapdl.mp("GYZ",  1, MAT1["GYZ"])
    mapdl.mp("PRXY", 1, MAT1["PRXY"])
    mapdl.mp("PRXZ", 1, MAT1["PRXZ"])
    mapdl.mp("PRYZ", 1, MAT1["PRYZ"])
    mapdl.mp("DENS", 1, MAT1["DENS"])
    print("Material 1 defined: Epoxy carbon woven prepreg")
    print(f"  E1={MAT1['EX']/1e9:.1f} GPa, E2={MAT1['EY']/1e9:.1f} GPa, "
          f"E3={MAT1['EZ']/1e9:.1f} GPa, rho={MAT1['DENS']:.0f} kg/m3")

    # --- Material 2: Honeycomb core ---
    mapdl.mp("EX",   2, MAT2["EX"])
    mapdl.mp("EY",   2, MAT2["EY"])
    mapdl.mp("EZ",   2, MAT2["EZ"])
    mapdl.mp("GXY",  2, MAT2["GXY"])
    mapdl.mp("GXZ",  2, MAT2["GXZ"])
    mapdl.mp("GYZ",  2, MAT2["GYZ"])
    mapdl.mp("PRXY", 2, MAT2["PRXY"])
    mapdl.mp("PRXZ", 2, MAT2["PRXZ"])
    mapdl.mp("PRYZ", 2, MAT2["PRYZ"])
    mapdl.mp("DENS", 2, MAT2["DENS"])
    print("Material 2 defined: Honeycomb core")
    print(f"  E1={MAT2['EX']/1e6:.1f} MPa, E2={MAT2['EY']/1e6:.1f} MPa, "
          f"E3={MAT2['EZ']/1e6:.1f} MPa, rho={MAT2['DENS']:.0f} kg/m3")

    # ------------------------------------------------------------------
    # 5.  DEFINE COMPOSITE SECTION (LAYERED SOLID)
    # ------------------------------------------------------------------
    print_banner("Step 5: Defining composite layered section")

    # SECTYPE: define section 1 as a solid section with sub-type "laye"
    # for layered composite.  The layered solid section tells SOLID186
    # how to distribute the plies through the element thickness.
    mapdl.sectype(1, "SOLID", "", "CompSandwich")

    # SECDATA: one call per ply layer.
    # Syntax: SECDATA, thickness, material, orientation, num_integ_pts
    # We use 3 integration points per ply for accuracy.
    total_thickness = 0.0
    for i, (mat_id, thickness, angle) in enumerate(LAYUP, start=1):
        mapdl.secdata(thickness, mat_id, angle, 3)
        total_thickness += thickness
        mat_name = "Carbon prepreg" if mat_id == 1 else "Honeycomb core"
        print(f"  Ply {i:2d}: {mat_name:20s}  t={thickness*1e3:7.3f} mm  "
              f"theta={angle:5.1f} deg")

    print(f"\n  Total laminate thickness: {total_thickness*1e3:.3f} mm")
    print(f"  Number of plies: {len(LAYUP)}")

    # ------------------------------------------------------------------
    # 6.  ASSIGN ATTRIBUTES AND MESH
    # ------------------------------------------------------------------
    print_banner("Step 6: Meshing with SOLID186 layered elements")

    # Assign material, element type, and section to all volumes.
    mapdl.allsel()
    mapdl.mat(1)           # default material (overridden by section layers)
    mapdl.type(1)          # element type 1 = SOLID186 layered
    mapdl.secnum(1)        # section 1 = our composite sandwich

    # Set mesh controls.
    if ELEMENT_SIZE > 0:
        mapdl.esize(ELEMENT_SIZE)
    else:
        # Use SmartSize to let MAPDL pick a reasonable mesh density.
        mapdl.smrtsize(4)  # level 4 = moderately fine

    # Volume mesh -- mesh all volumes with free meshing.
    mapdl.mshape(1, "3D")   # tetrahedral shape for free meshing
    mapdl.mshkey(0)          # free mesh
    mapdl.vmesh("ALL")

    num_nodes = mapdl.mesh.n_node
    num_elems = mapdl.mesh.n_elem
    print(f"Mesh generated: {num_nodes} nodes, {num_elems} elements")

    if num_elems == 0:
        print("ERROR: Meshing produced no elements. Check geometry/mesh settings.")
        mapdl.exit()
        sys.exit(1)

    print(f"  [{elapsed(t_start)}] Meshing complete.")

    # ------------------------------------------------------------------
    # 7.  BOUNDARY CONDITIONS -- FIXED SUPPORT AT HANDLE END
    # ------------------------------------------------------------------
    print_banner("Step 7: Applying fixed boundary conditions")

    # Strategy: identify the handle end by finding the extreme coordinate
    # of the model along its longest axis, then select nodes near that end.
    #
    # First, determine the bounding box of the model.
    mapdl.allsel()
    node_coords = mapdl.mesh.nodes   # (N x 3) array of node coordinates

    x_min, y_min, z_min = node_coords.min(axis=0)
    x_max, y_max, z_max = node_coords.max(axis=0)

    dx = x_max - x_min
    dy = y_max - y_min
    dz = z_max - z_min

    print(f"Model bounding box:")
    print(f"  X: {x_min*1e3:.2f} .. {x_max*1e3:.2f} mm  (span {dx*1e3:.2f} mm)")
    print(f"  Y: {y_min*1e3:.2f} .. {y_max*1e3:.2f} mm  (span {dy*1e3:.2f} mm)")
    print(f"  Z: {z_min*1e3:.2f} .. {z_max*1e3:.2f} mm  (span {dz*1e3:.2f} mm)")

    # The wrench handle runs along the longest dimension.
    # We fix one end (the end at the minimum coordinate of that axis).
    spans = {"X": dx, "Y": dy, "Z": dz}
    long_axis = max(spans, key=spans.get)
    long_span = spans[long_axis]

    print(f"\n  Longest axis: {long_axis} (span = {long_span*1e3:.2f} mm)")

    # Select nodes within 5% of the span from the minimum of the long axis.
    # This captures the handle end.
    tol = 0.05 * long_span  # 5% of the total span

    if long_axis == "X":
        loc_min = x_min
        mapdl.nsel("S", "LOC", "X", loc_min, loc_min + tol)
    elif long_axis == "Y":
        loc_min = y_min
        mapdl.nsel("S", "LOC", "Y", loc_min, loc_min + tol)
    else:
        loc_min = z_min
        mapdl.nsel("S", "LOC", "Z", loc_min, loc_min + tol)

    num_fixed = mapdl.mesh.n_node
    print(f"  Selected {num_fixed} nodes at handle end for fixed support")
    print(f"  ({long_axis} from {loc_min*1e3:.2f} to "
          f"{(loc_min + tol)*1e3:.2f} mm)")

    if num_fixed == 0:
        print("WARNING: No nodes selected for BC. Trying the opposite end ...")
        # Try the max end instead.
        if long_axis == "X":
            loc_max = x_max
            mapdl.nsel("S", "LOC", "X", loc_max - tol, loc_max)
        elif long_axis == "Y":
            loc_max = y_max
            mapdl.nsel("S", "LOC", "Y", loc_max - tol, loc_max)
        else:
            loc_max = z_max
            mapdl.nsel("S", "LOC", "Z", loc_max - tol, loc_max)
        num_fixed = mapdl.mesh.n_node
        print(f"  Re-selected {num_fixed} nodes at opposite end")

    # Apply fixed constraint (all DOFs = 0) to selected nodes.
    mapdl.d("ALL", "ALL", 0)
    print("  Fixed support applied (UX=UY=UZ=0 on selected nodes)")

    # Re-select everything for the solve.
    mapdl.allsel()
    print(f"  [{elapsed(t_start)}] Boundary conditions applied.")

    # ------------------------------------------------------------------
    # 8.  MODAL ANALYSIS
    # ------------------------------------------------------------------
    print_banner("Step 8: Modal analysis (Block Lanczos)")

    mapdl.run("/SOLU")
    mapdl.antype("MODAL")                # Analysis type = modal
    mapdl.modopt("LANB", NUM_MODES,      # Block Lanczos solver
                 FREQ_START, FREQ_END)   # Frequency range
    mapdl.eqslv("SPARSE")               # Sparse direct solver
    mapdl.mxpand(NUM_MODES, 0, 0, "YES")  # Expand all modes, compute stresses

    print(f"  Solver      : Block Lanczos")
    print(f"  Modes       : {NUM_MODES}")
    print(f"  Freq range  : {FREQ_START:.0f} - {FREQ_END:.0f} Hz")
    print(f"  Mode expand : YES (stresses computed)")

    print("\n  Solving modal analysis ...")
    t_modal = time.time()
    mapdl.solve()
    mapdl.finish()
    print(f"  Modal solve completed in {time.time() - t_modal:.1f} seconds.")

    # ------------------------------------------------------------------
    # 8b. POST-PROCESS MODAL RESULTS (quick summary)
    # ------------------------------------------------------------------
    print_banner("Step 8b: Modal results summary")

    mapdl.post1()
    # Read the natural frequencies from the result file.
    # SET,LIST prints a table of all load steps / substeps (i.e. modes).
    modal_summary = mapdl.set("LIST")
    print(modal_summary)

    # Also extract frequencies into a Python list for later reference.
    nat_freqs = []
    for mode_num in range(1, NUM_MODES + 1):
        try:
            mapdl.set(1, mode_num)
            freq = mapdl.get("FREQ_VAL", "ACTIVE", 0, "SET", "FREQ")
            nat_freqs.append(freq)
            print(f"  Mode {mode_num:3d}:  {freq:10.3f} Hz")
        except Exception:
            # If fewer modes were found than requested, stop gracefully.
            break

    num_found = len(nat_freqs)
    print(f"\n  Found {num_found} modes in range "
          f"{FREQ_START:.0f}-{FREQ_END:.0f} Hz")

    if num_found == 0:
        print("ERROR: No modes found. Check BCs and frequency range.")
        mapdl.exit()
        sys.exit(1)

    mapdl.finish()
    print(f"  [{elapsed(t_start)}] Modal analysis complete.")

    # ------------------------------------------------------------------
    # 9.  SPECTRUM (PSD) ANALYSIS
    # ------------------------------------------------------------------
    print_banner("Step 9: PSD spectrum analysis")

    mapdl.run("/SOLU")
    mapdl.antype("SPECTR")               # Analysis type = spectrum
    mapdl.spopt("PSD", NUM_MODES, "ON")  # PSD with NUM_MODES modes,
                                          # "ON" = compute element results

    # Define PSD excitation type.
    # PSDUNIT: table 1 is acceleration in units of G (gravitational g).
    mapdl.psdunit(1, "ACCG", 9.80665)    # ACCG = acceleration in G;
                                          # the value 9.80665 converts G -> m/s^2

    # Build the PSD frequency/value table.
    # PSDFRQ: define frequencies for PSD table 1.
    # PSDVAL: define corresponding PSD values.
    freqs = [f for f, _ in PSD_TABLE]
    vals  = [v for _, v in PSD_TABLE]

    print("  PSD input table (G^2/Hz):")
    for f, v in PSD_TABLE:
        print(f"    {f:8.1f} Hz  :  {v:.4f} G^2/Hz")

    # APDL PSDFRQ and PSDVAL accept up to 6 values per command call.
    # With only 4 points we can do it in one call each.
    # PSDFRQ, table_num, mode_type(0=all), freq1, freq2, ...
    mapdl.psdfrq(1, 1, freqs[0], freqs[1], freqs[2], freqs[3])
    mapdl.psdval(1, vals[0], vals[1], vals[2], vals[3])

    # Log-log interpolation is typical for PSD data.
    # PSDLOG: set interpolation to log-log for both frequency and PSD axes.
    # Not all versions support PSDLOG directly; use PSDSPL if needed.
    # Default in MAPDL is log-log for PSD, so this may be redundant but
    # explicit is better than implicit.

    # Apply base excitation as a boundary condition.
    # The PSD is applied in the Y-direction at all fixed DOFs.
    # D command with PSDISP label ties the PSD table to those DOFs.
    # First re-select the handle-end nodes that have fixed BCs.
    if long_axis == "X":
        mapdl.nsel("S", "LOC", "X", loc_min, loc_min + tol)
    elif long_axis == "Y":
        mapdl.nsel("S", "LOC", "Y", loc_min, loc_min + tol)
    else:
        mapdl.nsel("S", "LOC", "Z", loc_min, loc_min + tol)

    # For base excitation PSD in Y direction, we use:
    #   D, ALL, UY, 1.0   (unit displacement spectrum reference)
    # with the PSD table already defined.
    # The base excitation is applied through the constraint DOFs.
    mapdl.d("ALL", "UY", 1.0)

    # Re-select all for the solve.
    mapdl.allsel()

    # Participation factor: this links PSD table 1 to the base excitation
    # in the Y direction for the nodes with D constraints.
    # PFACT, table_num, type (BASE = base excitation), direction
    mapdl.pfact(1, "BASE")

    # Specify what results to combine.
    # PSDRES: controls which 1-sigma results are written.
    #   PSDRES, label, rel_key
    #   label = DISP -> displacement results
    #   label = VELO -> velocity results
    #   label = ACEL -> acceleration results
    #   label = STRE -> stress results
    #   rel_key = 1 -> relative to base, 2 -> absolute
    # We request both displacement (relative) and stress results.
    mapdl.psdres("DISP", 1)   # 1-sigma relative displacement
    mapdl.psdres("VELO", 1)   # 1-sigma relative velocity
    mapdl.psdres("ACEL", 1)   # 1-sigma relative acceleration
    mapdl.psdres("STRE", 1)   # 1-sigma stress

    # Set damping -- constant modal damping ratio.
    # A typical value for composite structures is 1-3%.
    mapdl.dmprat(0.02)  # 2% critical damping for all modes

    print(f"  Damping ratio: 2% (constant for all modes)")
    print(f"  Base excitation direction: UY")
    print(f"  Results requested: DISP, VELO, ACEL, STRE (1-sigma relative)")

    # Solve the PSD analysis.
    print("\n  Solving PSD analysis ...")
    t_psd = time.time()
    mapdl.solve()
    mapdl.finish()
    print(f"  PSD solve completed in {time.time() - t_psd:.1f} seconds.")
    print(f"  [{elapsed(t_start)}] Spectrum analysis complete.")

    # ------------------------------------------------------------------
    # 10. POST-PROCESSING -- 1-SIGMA RESULTS
    # ------------------------------------------------------------------
    print_banner("Step 10: Post-processing PSD results")

    mapdl.post1()

    # The PSD analysis stores 1-sigma results in specific load-step/substep
    # combinations.  For a single PSD table:
    #   Load step 1 = modal solution (not of interest here)
    #   Load step 2 = PSD results
    # Sub-step 1 within the PSD step contains the 1-sigma combination.
    #
    # We use SET to navigate to the correct result set.
    # MAPDL convention for PSD post-processing:
    #   SET,2,1 -> 1-sigma displacement/stress from PSD
    # If that does not work, iterate through available sets.

    # List available result sets to find the PSD 1-sigma set.
    set_list = mapdl.set("LIST")
    print("Available result sets:")
    print(set_list)

    # Attempt to read the last available result set, which is typically
    # the 1-sigma combined PSD result.
    try:
        mapdl.set("LAST")
    except Exception as e:
        print(f"  Warning setting result set: {e}")
        # Fall back to set 2,1
        try:
            mapdl.set(2, 1)
        except Exception:
            mapdl.set(1, 1)

    # ---- 1-sigma Displacement Results ----
    print("\n--- 1-Sigma Displacement Results ---")

    # Get the maximum displacement magnitude.
    # PLNSOL plots nodal solution; PRNSOL prints it.
    # We can extract via mapdl.post_processing for PyMAPDL convenience.
    try:
        # Try PyMAPDL post-processing interface
        disp_y = mapdl.post_processing.nodal_displacement("Y")
        disp_mag = mapdl.post_processing.nodal_displacement("NORM")

        max_uy     = np.max(np.abs(disp_y))
        max_uy_idx = np.argmax(np.abs(disp_y))
        max_disp   = np.max(disp_mag)

        print(f"  Max |UY| (1-sigma) = {max_uy*1e6:.4f} um  "
              f"({max_uy*1e3:.6f} mm)")
        print(f"  Max |U|  (1-sigma) = {max_disp*1e6:.4f} um  "
              f"({max_disp*1e3:.6f} mm)")
    except Exception as e:
        print(f"  Could not extract displacements via Python API: {e}")
        print("  Falling back to APDL PRNSOL ...")
        mapdl.run("PRNSOL,U,COMP")

    # ---- Plot 1-sigma Displacement (Y-component) ----
    print("\n  Plotting 1-sigma UY displacement ...")
    try:
        mapdl.post_processing.plot_nodal_displacement(
            "Y",
            title="1-Sigma UY Displacement (PSD)",
            show_node_numbering=False,
            cpos="iso",
            screenshot="psd_disp_uy.png",
            off_screen=True,
        )
        print("  Saved: psd_disp_uy.png")
    except Exception as e:
        print(f"  Plot via Python failed ({e}); using MAPDL plotter ...")
        try:
            mapdl.run("PLNSOL,U,Y,0,1")
        except Exception:
            print("  MAPDL plot also failed. Skipping displacement plot.")

    # ---- Plot 1-sigma Displacement magnitude ----
    print("  Plotting 1-sigma displacement magnitude ...")
    try:
        mapdl.post_processing.plot_nodal_displacement(
            "NORM",
            title="1-Sigma Displacement Magnitude (PSD)",
            show_node_numbering=False,
            cpos="iso",
            screenshot="psd_disp_mag.png",
            off_screen=True,
        )
        print("  Saved: psd_disp_mag.png")
    except Exception as e:
        print(f"  Plot via Python failed ({e}); using MAPDL plotter ...")
        try:
            mapdl.run("PLNSOL,U,SUM,0,1")
        except Exception:
            print("  MAPDL plot also failed. Skipping magnitude plot.")

    # ---- 1-sigma Stress Results ----
    print("\n--- 1-Sigma Stress Results ---")

    try:
        # Von Mises equivalent stress
        stress_eqv = mapdl.post_processing.nodal_eqv_stress()
        max_seqv = np.max(stress_eqv)
        print(f"  Max von Mises stress (1-sigma) = {max_seqv/1e6:.4f} MPa")

        # Component stresses
        stress_x = mapdl.post_processing.nodal_component_stress("X")
        stress_y = mapdl.post_processing.nodal_component_stress("Y")
        stress_z = mapdl.post_processing.nodal_component_stress("Z")
        print(f"  Max |SX| (1-sigma) = {np.max(np.abs(stress_x))/1e6:.4f} MPa")
        print(f"  Max |SY| (1-sigma) = {np.max(np.abs(stress_y))/1e6:.4f} MPa")
        print(f"  Max |SZ| (1-sigma) = {np.max(np.abs(stress_z))/1e6:.4f} MPa")
    except Exception as e:
        print(f"  Could not extract stresses via Python API: {e}")
        print("  Falling back to APDL PRNSOL ...")
        mapdl.run("PRNSOL,S,COMP")

    # ---- Plot 1-sigma von Mises Stress ----
    print("\n  Plotting 1-sigma von Mises stress ...")
    try:
        mapdl.post_processing.plot_nodal_eqv_stress(
            title="1-Sigma von Mises Stress (PSD)",
            show_node_numbering=False,
            cpos="iso",
            screenshot="psd_stress_eqv.png",
            off_screen=True,
        )
        print("  Saved: psd_stress_eqv.png")
    except Exception as e:
        print(f"  Plot via Python failed ({e}); using MAPDL plotter ...")
        try:
            mapdl.run("PLNSOL,S,EQV,0,1")
        except Exception:
            print("  MAPDL plot also failed. Skipping stress plot.")

    # ---- Summary Table ----
    print_banner("ANALYSIS SUMMARY")
    print(f"  Geometry file      : {PARASOLID_FILE}")
    print(f"  Element type       : SOLID186 layered (KEYOPT(3)=1)")
    print(f"  Mesh               : {num_elems} elements, {num_nodes} nodes")
    print(f"  Composite layup    : [0/0/45/45/90/core/90/45/45/0/0]")
    print(f"  Total thickness    : {total_thickness*1e3:.3f} mm")
    print(f"  Boundary condition : Fixed at handle end ({long_axis} min)")
    print(f"  Modal modes found  : {num_found}")
    if num_found > 0:
        print(f"  First natural freq : {nat_freqs[0]:.3f} Hz")
        if num_found > 1:
            print(f"  Last natural freq  : {nat_freqs[-1]:.3f} Hz")
    print(f"  PSD excitation     : Base accel. in Y, {PSD_TABLE[0][0]:.0f}"
          f"-{PSD_TABLE[-1][0]:.0f} Hz")
    print(f"  Damping            : 2% constant modal damping")
    try:
        print(f"  Max 1-sigma |UY|   : {max_uy*1e6:.4f} um")
        print(f"  Max 1-sigma |U|    : {max_disp*1e6:.4f} um")
    except NameError:
        pass
    try:
        print(f"  Max 1-sigma SEQV   : {max_seqv/1e6:.4f} MPa")
    except NameError:
        pass

    print(f"\n  Total wall time    : {elapsed(t_start)}")

    # ------------------------------------------------------------------
    # 11. CLEAN UP
    # ------------------------------------------------------------------
    print_banner("Step 11: Cleanup")

    # Save the database before exiting.
    mapdl.finish()
    mapdl.run("/SAVE,ALL")
    print("  Database saved.")

    # Exit MAPDL session.
    mapdl.exit()
    print("  MAPDL session closed.")
    print(f"\n  Analysis completed successfully. [{elapsed(t_start)}]")


# ===========================================================================
#                       ENTRY POINT
# ===========================================================================
if __name__ == "__main__":
    try:
        main()
    except Exception as exc:
        print(f"\n{'!'*60}")
        print(f"  FATAL ERROR: {exc}")
        print(f"{'!'*60}")
        import traceback
        traceback.print_exc()
        sys.exit(1)
