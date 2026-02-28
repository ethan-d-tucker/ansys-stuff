# Composite Random Vibration Analysis - Setup Guide

## Prerequisites

- **Ansys Student 2025** installed
- **Python 3.9+** (3.10 or 3.11 recommended)

## Installation

```bash
# Create a virtual environment (recommended)
python -m venv ansys_env
ansys_env\Scripts\activate

# Install PyMAPDL
pip install ansys-mapdl-core

# Install visualization (optional but recommended)
pip install pyvista matplotlib
```

## Ansys Student License Note

The Student version limits mesh size to ~128k nodes/elements. If your wrench
geometry produces a mesh exceeding this, increase `ELEMENT_SIZE` in the script
or simplify the geometry.

## Running the Analysis

```bash
python composite_random_vibration.py
```

The script will:
1. Launch an MAPDL session automatically
2. Import `heavyDutyWrench.x_t` from your Downloads folder
3. Define composite materials and layered section
4. Mesh with SOLID186 layered elements
5. Fix the handle end
6. Run modal analysis (20 modes, 0-3000 Hz)
7. Run PSD random vibration (Y-axis base excitation)
8. Print 1-sigma results and save contour plots

## Output Files

- `psd_disp_uy.png` — 1-sigma Y displacement contour
- `psd_disp_mag.png` — 1-sigma displacement magnitude contour
- `psd_stress_eqv.png` — 1-sigma von Mises stress contour
- Console output with full summary table

## Configuration

All parameters are at the top of `composite_random_vibration.py`:

| Parameter | Default | Description |
|---|---|---|
| `PARASOLID_FILE` | Downloads/heavyDutyWrench.x_t | Geometry file path |
| `PLY_THICKNESS_CARBON` | 0.286 mm | Carbon prepreg ply thickness |
| `CORE_THICKNESS` | 6.35 mm | Honeycomb core thickness |
| `NUM_MODES` | 20 | Number of modes to extract |
| `FREQ_END` | 3000 Hz | Upper frequency bound |
| `PSD_TABLE` | MIL-STD-810-ish profile | PSD input curve |
| `ELEMENT_SIZE` | 0 (auto) | Mesh element size |

## Troubleshooting

- **"Cannot find Ansys installation"**: Set the path manually:
  ```python
  mapdl = launch_mapdl(exec_file=r"C:\Program Files\ANSYS Inc\v251\ansys\bin\winx64\ANSYS251.exe")
  ```
- **Mesh too large for Student license**: Increase `ELEMENT_SIZE` (e.g., `0.005` for 5mm elements)
- **No modes found**: Check boundary conditions — model may be free-floating
