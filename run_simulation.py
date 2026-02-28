"""
Composite Wrench Random Vibration (PSD) Analysis
=================================================
Full end-to-end simulation using WrenchParasolid.x_t geometry.

Uses validated methods from debug sessions:
  - Parasolid import via ac4para.exe with P_SCHEMA
  - mapdl.input() for ANF loading (not run("/INPUT,..."))
  - SOLID187 (10-node tet) free mesh with orthotropic material
  - Modal analysis (Block Lanczos, 20 modes, 0-3000 Hz)
  - PSD 1-sigma computed manually from modal data
    (Student edition doesn't store PSD combination results in .rst)
"""
import sys
import os
import time
import subprocess
import shutil
import numpy as np

from ansys.mapdl.core import launch_mapdl

t0 = time.time()

def elapsed():
    dt = time.time() - t0
    m, s = divmod(dt, 60)
    return f"{int(m):02d}:{s:05.2f}"

script_dir = os.path.dirname(os.path.abspath(__file__))
ansys_base = r"C:\Program Files\ANSYS Inc\ANSYS Student\v252"

# ===================================================================
# CONFIGURATION
# ===================================================================
PARASOLID_FILE = os.path.join(script_dir, "WrenchParasolid.x_t")
ELEMENT_SIZE = 0.003  # 3mm for good resolution

# Material: Epoxy carbon woven prepreg (orthotropic)
MAT1 = {
    "EX": 60.0e9, "EY": 60.0e9, "EZ": 10.0e9,
    "GXY": 5.0e9, "GXZ": 4.0e9, "GYZ": 4.0e9,
    "PRXY": 0.04, "PRXZ": 0.30, "PRYZ": 0.30,
    "DENS": 1420.0,
}

# PSD input (G^2/Hz)
PSD_TABLE = [(20.0, 0.010), (80.0, 0.040), (350.0, 0.040), (2000.0, 0.007)]
DAMPING_RATIO = 0.02  # 2% critical damping

NUM_MODES = 20
FREQ_START = 0.0
FREQ_END = 3000.0


def compute_psd_manual(nat_freqs, mode_shapes, psd_table, damping):
    """
    Compute 1-sigma PSD response manually from modal data via SRSS.

    For each mode i with natural frequency f_i:
      H_i(f) = 1 / sqrt((1-(f/f_i)^2)^2 + (2*zeta*f/f_i)^2)
      sigma_i^2 = integral[ |H_i(f)|^2 * S_a(f) ] df / (2*pi*f_i)^4

    1-sigma displacement at node j:
      sigma_j = sqrt( sum_i (phi_ji^2 * sigma_i^2) )
    """
    g = 9.80665
    n_modes = len(nat_freqs)

    psd_freqs = np.array([f for f, _ in psd_table])
    psd_vals = np.array([v for _, v in psd_table]) * g**2  # -> (m/s^2)^2/Hz

    f_min = max(psd_freqs[0], 1.0)
    f_max = psd_freqs[-1]
    f_grid = np.logspace(np.log10(f_min), np.log10(f_max), 2000)

    log_psd = np.interp(np.log10(f_grid), np.log10(psd_freqs),
                        np.log10(psd_vals))
    psd_interp = 10**log_psd

    modal_sigma2 = np.zeros(n_modes)

    print("  Modal PSD contributions:")
    for i, fi in enumerate(nat_freqs):
        omega_i = 2 * np.pi * fi
        r = f_grid / fi
        H2 = 1.0 / ((1 - r**2)**2 + (2 * damping * r)**2)
        modal_sigma2[i] = np.trapezoid(H2 * psd_interp, f_grid) / omega_i**4
        print(f"    Mode {i+1:2d} ({fi:8.2f} Hz): "
              f"RMS modal disp = {np.sqrt(modal_sigma2[i]):.6e} m")

    results = {}
    for comp, shapes in mode_shapes.items():
        sigma2_nodes = np.zeros(shapes.shape[1])
        for i in range(n_modes):
            sigma2_nodes += shapes[i, :]**2 * modal_sigma2[i]
        results[comp] = np.sqrt(sigma2_nodes)

    return results, modal_sigma2


# ===================================================================
# STEP 1: Launch MAPDL
# ===================================================================
print("=" * 60)
print("  WRENCH PSD ANALYSIS (WrenchParasolid.x_t)")
print("=" * 60)
print()
print(f"[{elapsed()}] Step 1: Launching MAPDL...")

mapdl = launch_mapdl(override=True, loglevel="WARNING", start_timeout=120)
print(f"  MAPDL v{mapdl.version}")
work_dir = mapdl.directory

# ===================================================================
# STEP 2: Import Parasolid Geometry
# ===================================================================
print(f"[{elapsed()}] Step 2: Importing Parasolid geometry...")

converter = os.path.join(ansys_base, "ansys", "ac4", "bin", "para",
                         "winx64", "ac4para.exe")
schema_dir = os.path.join(ansys_base, "commonfiles", "CAD", "Siemens",
                          "Parasolid36.1.227", "winx64", "schema")

shutil.copy2(PARASOLID_FILE, os.path.join(work_dir, "WrenchParasolid.x_t"))

env = os.environ.copy()
env["PATH"] = (os.path.join(ansys_base, "ansys", "bin", "winx64") + ";" +
               os.path.dirname(converter) + ";" +
               os.path.join(ansys_base, "commonfiles", "CAD", "Siemens",
                            "Parasolid36.1.227", "winx64") + ";" +
               env.get("PATH", ""))
env["P_SCHEMA"] = schema_dir

result = subprocess.run(
    [converter, "WrenchParasolid.x_t", "Wrench.anf", "SOLIDS", "ANF"],
    cwd=work_dir, env=env, timeout=120,
    stdout=subprocess.PIPE, stderr=subprocess.PIPE
)
if result.returncode != 0:
    print(f"  FATAL: ac4 converter failed: {result.stdout.decode()}")
    sys.exit(1)

anf_file = os.path.join(work_dir, "Wrench.anf")
print(f"  ac4 converter: {os.path.getsize(anf_file)} bytes")

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
print(f"  Geometry: {nv} vol, {na} areas")

# ===================================================================
# STEP 3: Element Type + Material
# ===================================================================
print(f"[{elapsed()}] Step 3: Element type + material...")

mapdl.et(1, "SOLID187")
for prop, val in MAT1.items():
    mapdl.mp(prop, 1, val)
print(f"  SOLID187, Carbon/Epoxy (E={MAT1['EX']/1e9:.0f} GPa)")

# ===================================================================
# STEP 4: Mesh
# ===================================================================
print(f"[{elapsed()}] Step 4: Meshing (esize={ELEMENT_SIZE*1000:.1f}mm)...")

mapdl.allsel()
mapdl.mat(1)
mapdl.type(1)
mapdl.esize(ELEMENT_SIZE)
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

for i, ax in enumerate(["X", "Y", "Z"]):
    print(f"  {ax}: {mins[i]*1e3:.1f} to {maxs[i]*1e3:.1f} mm "
          f"(span {spans[i]*1e3:.1f})")

# ===================================================================
# STEP 5: Boundary Conditions
# ===================================================================
print(f"[{elapsed()}] Step 5: Boundary conditions...")

longest_idx = np.argmax(spans)
axis_names = ["X", "Y", "Z"]
long_axis = axis_names[longest_idx]
tol = 0.05 * spans[longest_idx]
fix_min = mins[longest_idx]

mapdl.nsel("S", "LOC", long_axis, fix_min, fix_min + tol)
n_fixed = mapdl.mesh.n_node
mapdl.d("ALL", "ALL", 0)
mapdl.allsel()
print(f"  Fixed {n_fixed} nodes at {long_axis} handle end")

# ===================================================================
# STEP 6: Modal Analysis
# ===================================================================
print(f"[{elapsed()}] Step 6: Modal analysis ({NUM_MODES} modes, "
      f"{FREQ_START:.0f}-{FREQ_END:.0f} Hz)...")

mapdl.run("/SOLU")
mapdl.antype("MODAL")
mapdl.modopt("LANB", NUM_MODES, FREQ_START, FREQ_END)
mapdl.eqslv("SPARSE")
mapdl.mxpand(NUM_MODES, 0, 0, "YES")

t_modal = time.time()
mapdl.solve()
mapdl.finish()
dt_modal = time.time() - t_modal
print(f"  Solve time: {dt_modal:.1f}s")

# ===================================================================
# STEP 7: Extract Modal Data
# ===================================================================
print(f"[{elapsed()}] Step 7: Extracting modal data...")

mapdl.post1()

nat_freqs = []
mode_shapes_y = []
mode_shapes_x = []
mode_shapes_z = []
mode_shapes_mag = []
stress_eqv_modes = []

for mode in range(1, NUM_MODES + 1):
    try:
        mapdl.set(1, mode)
        freq = mapdl.get("FREQ_VAL", "ACTIVE", 0, "SET", "FREQ")
        nat_freqs.append(freq)

        dy = mapdl.post_processing.nodal_displacement("Y")
        dx = mapdl.post_processing.nodal_displacement("X")
        dz = mapdl.post_processing.nodal_displacement("Z")

        mode_shapes_y.append(dy.copy())
        mode_shapes_x.append(dx.copy())
        mode_shapes_z.append(dz.copy())

        try:
            seqv = mapdl.post_processing.nodal_eqv_stress()
            stress_eqv_modes.append(seqv.copy())
        except Exception:
            stress_eqv_modes.append(np.zeros_like(dy))

        print(f"  Mode {mode:2d}: {freq:10.2f} Hz  "
              f"max|UY|={np.max(np.abs(dy)):.4f}")
    except Exception:
        break

n_modes = len(nat_freqs)
nat_freqs = np.array(nat_freqs)
mapdl.finish()

# ===================================================================
# STEP 8: Compute PSD 1-Sigma Response
# ===================================================================
print(f"\n[{elapsed()}] Step 8: Computing PSD 1-sigma response...")

mode_shapes = {
    "Y": np.array(mode_shapes_y),
    "X": np.array(mode_shapes_x),
    "Z": np.array(mode_shapes_z),
}

psd_results, modal_sigma2 = compute_psd_manual(
    nat_freqs, mode_shapes, PSD_TABLE, DAMPING_RATIO
)

# Stress via SRSS
stress_shapes = np.array(stress_eqv_modes)
sigma2_stress = np.zeros(stress_shapes.shape[1])
for i in range(n_modes):
    sigma2_stress += stress_shapes[i, :]**2 * modal_sigma2[i]
sigma_stress = np.sqrt(sigma2_stress)

# Displacement magnitude via SRSS of components
sigma_ux = psd_results["X"]
sigma_uy = psd_results["Y"]
sigma_uz = psd_results["Z"]
sigma_umag = np.sqrt(sigma_ux**2 + sigma_uy**2 + sigma_uz**2)

# Identify free nodes
all_nids = mapdl.mesh.nnum
mapdl.prep7()
mapdl.nsel("S", "LOC", long_axis, fix_min, fix_min + tol)
fixed_set = set(mapdl.mesh.nnum)
mapdl.allsel()
free_mask = np.array([n not in fixed_set for n in all_nids])

# Grms of input PSD
g = 9.80665
psd_freqs_arr = np.array([f for f, _ in PSD_TABLE])
psd_vals_arr = np.array([v for _, v in PSD_TABLE])
f_grid = np.logspace(np.log10(max(psd_freqs_arr[0], 1.0)),
                     np.log10(psd_freqs_arr[-1]), 2000)
log_psd = np.interp(np.log10(f_grid), np.log10(psd_freqs_arr),
                    np.log10(psd_vals_arr))
input_grms = np.sqrt(np.trapezoid(10**log_psd, f_grid))

# ===================================================================
# RESULTS
# ===================================================================
print()
print("=" * 60)
print("  RESULTS")
print("=" * 60)

print()
print("  --- Natural Frequencies ---")
for i, f in enumerate(nat_freqs, 1):
    print(f"    Mode {i:2d}: {f:10.2f} Hz")

print()
print("  --- 1-Sigma Displacement (PSD Response) ---")
for comp, vals in [("UX", sigma_ux), ("UY", sigma_uy),
                   ("UZ", sigma_uz), ("|U|", sigma_umag)]:
    max_val = np.max(vals[free_mask])
    max_idx = np.argmax(vals[free_mask])
    actual_idx = np.where(free_mask)[0][max_idx]
    node_id = all_nids[actual_idx]
    pos = coords[actual_idx]
    print(f"    Max {comp:4s} = {max_val*1e6:.4f} um  "
          f"({max_val*1e3:.6f} mm) @ node {node_id}")

print()
print("  --- 1-Sigma von Mises Stress ---")
max_s = np.max(sigma_stress[free_mask])
max_s_idx = np.argmax(sigma_stress[free_mask])
actual_idx = np.where(free_mask)[0][max_s_idx]
node_id = all_nids[actual_idx]
pos = coords[actual_idx]
print(f"    Max SEQV = {max_s/1e6:.4f} MPa @ node {node_id} "
      f"({pos[0]*1e3:.1f}, {pos[1]*1e3:.1f}, {pos[2]*1e3:.1f}) mm")

print()
print(f"  Input PSD Grms = {input_grms:.4f} G")

# ===================================================================
# SUMMARY
# ===================================================================
print()
print("=" * 60)
print("  ANALYSIS SUMMARY")
print("=" * 60)
print(f"  Geometry        : {os.path.basename(PARASOLID_FILE)}")
print(f"  Element type    : SOLID187 (10-node tet)")
print(f"  Mesh            : {ne} elements, {nn} nodes")
print(f"  Material        : Carbon/Epoxy orthotropic")
print(f"                    E1=E2={MAT1['EX']/1e9:.0f} GPa, "
      f"E3={MAT1['EZ']/1e9:.0f} GPa, rho={MAT1['DENS']:.0f} kg/m3")
print(f"  BC              : Fixed at handle end ({long_axis} min)")
print(f"  Modes found     : {n_modes}")
print(f"  Freq range      : {nat_freqs[0]:.2f} - {nat_freqs[-1]:.2f} Hz")
print(f"  PSD excitation  : {PSD_TABLE[0][0]:.0f}-{PSD_TABLE[-1][0]:.0f} Hz,"
      f" Y-dir, {input_grms:.4f} Grms")
print(f"  Damping         : {DAMPING_RATIO*100:.0f}% constant modal")
print(f"  Max 1s |UY|     : {np.max(sigma_uy[free_mask])*1e6:.4f} um")
print(f"  Max 1s |U|      : {np.max(sigma_umag[free_mask])*1e6:.4f} um")
print(f"  Max 1s SEQV     : {max_s/1e6:.4f} MPa")
print(f"  Total time      : {elapsed()}")

# Cleanup
mapdl.finish()
mapdl.exit()
print(f"\n[OK] Simulation complete! [{elapsed()}]")
