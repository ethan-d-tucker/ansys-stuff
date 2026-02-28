# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ANSYS composite random vibration (PSD) analysis project. Supports dynamic part selection and MIL-STD-810H multi-environment qualification with composite failure analysis.

### Primary workflow (MIL-STD-810H)
1. **`run_milstd_analysis.py`** — **Main entry point.** Multi-environment PSD analysis per MIL-STD-810H Method 514.8. Modal solve once, PSD for each environment x 3 axes, composite failure analysis (Tsai-Wu + Max Stress + FoS), generates professional DOCX report. Accepts any Parasolid/STEP geometry.
2. **`ansys_mcp_server.py`** — FastMCP server exposing 35+ ANSYS MAPDL tools for interactive Claude-driven FEA. Includes MIL-STD tools: `get_milstd_profiles()`, `get_material_library()`, `get_default_layup()`, `run_milstd_psd_analysis()`, `compute_composite_failure_standalone()`.

### Supporting modules
3. **`simulation_engine.py`** — Parameterised simulation core. `SimulationConfig` dataclass, ac4 Parasolid import, SOLID187 mesh, modal solve, PSD SRSS with all 6 stress components. `run_multi_environment()` for batch runs.
4. **`mil_std_profiles.py`** — MIL-STD-810H Method 514.8 vibration environment definitions (4 profiles: General Minimum Integrity, Truck Transport, Helicopter, Jet Aircraft).
5. **`material_library.py`** — Material database with elastic properties (for ANSYS) and strength allowables (for failure analysis). Default 21-ply all-carbon symmetric laminate.
6. **`composite_failure.py`** — Tsai-Wu and Max Stress failure analysis engine. Ply-by-ply stress rotation, FoS computation, core failure check.
7. **`milstd_report.py`** — MIL-STD-810H tailored DOCX report generator with requirements traceability, failure assessment, compliance matrix.

### Legacy / reference scripts
8. **`run_simulation.py`** — Original single-config PSD analysis (hardcoded wrench). Still works standalone.
9. **`run_and_report.py`** — Original simulation + DOCX report (hardcoded wrench, ~30 images).
10. **`composite_random_vibration.py`** — Legacy script (broken ANF/SOLID186 approach, do not use).

Supporting files: `generate_plots.py` (matplotlib visualization), `generate_report.py` (HTML report generator), `psd_curve.csv` (input spectrum), `WrenchParasolid.x_t` (primary CAD geometry), `heavyDutyWrench.iges` (IGES fallback), `anf_commands.json` (ANF geometry fallback — has broken topology, avoid).

## Environment

- **Python**: `C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe` (system install, no venv)
- **ANSYS**: Student 2025 R2 (v25.2) at `C:\Program Files\ANSYS Inc\ANSYS Student\v252`
- **Packages**: ansys-mapdl-core, fastmcp, matplotlib, pyvista, numpy (v2.x), python-docx
- **Shell**: Git Bash on Windows — use `//` for flag prefixes in taskkill (e.g., `taskkill //F //IM ANSYS252.exe`)
- Student license limits mesh to ~128k nodes/elements

## Running

```bash
# MIL-STD-810H multi-environment analysis (recommended)
"C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe" run_milstd_analysis.py \
    --geometry WrenchParasolid.x_t \
    --part-name "Heavy-Duty Wrench" \
    --profiles MIN_INTEGRITY,HELICOPTER,JET_AIRCRAFT
# Output: report_output/MIL_STD_810H_PSD_Report.docx

# All profiles: MIN_INTEGRITY, TRUCK_TRANSPORT, HELICOPTER, JET_AIRCRAFT
# Options: --element-size 0.003 --damping 0.02 --required-fos 1.5

# Legacy: single-config simulation
"C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe" run_simulation.py

# Legacy: single-config + DOCX report
"C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe" run_and_report.py

# MCP server is auto-launched by Claude Code via .mcp.json
```

## Interactive Workflow (via Claude MCP)

1. User specifies geometry file and part name
2. Claude calls `get_default_layup()` and presents the composite layup for review
3. User confirms or requests modifications to the layup
4. Claude calls `get_milstd_profiles()` and presents available MIL-STD environments
5. User selects which profiles to test
6. Claude calls `run_milstd_psd_analysis()` — runs modal solve once, PSD per env x 3 axes
7. Report generated automatically with failure analysis, FoS, compliance matrix

## Architecture

### MIL-STD-810H Analysis (`run_milstd_analysis.py`)
- Accepts any Parasolid (.x_t) or STEP (.stp) geometry via `--geometry`
- Runs modal solve ONCE, then PSD SRSS for each environment x each axis (X, Y, Z)
- Extracts all 6 stress components per mode (SX, SY, SZ, SXY, SXZ, SYZ) for failure analysis
- Modal participation factors weight each mode's response by excitation direction (direction-dependent PSD)
- Composite failure via `composite_failure.py`: Tsai-Wu + Max Stress criteria, ply-by-ply stress rotation
- Factors of safety computed per node per ply (free nodes only — excludes BC singularities); worst (critical) ply identified
- Core failure check runs only when layup contains core material
- DOCX report via `milstd_report.py`: 9 sections, requirements traceability matrix, compliance matrix
- 4 MIL-STD profiles available: MIN_INTEGRITY (REQ-VIB-001), TRUCK_TRANSPORT (REQ-VIB-002), HELICOPTER (REQ-VIB-003), JET_AIRCRAFT (REQ-VIB-004)
- Typical run: ~10-20s depending on mesh density and number of profiles

### Working Simulation (`run_simulation.py`)
- Imports Parasolid geometry via `ac4para.exe` converter with `P_SCHEMA` env var
- Uses `mapdl.input()` for ANF loading (NOT `mapdl.run("/INPUT,...")`)
- SOLID187 (10-node tet) with free mesh and orthotropic carbon/epoxy properties
- Modal analysis (Block Lanczos, 6 modes found in 0-3000 Hz range)
- PSD 1-sigma computed manually from modal data via SRSS (Student edition workaround)
- Typical run: ~5 seconds total

### Report Generator (`run_and_report.py`)
- Runs full simulation (same as `run_simulation.py`) then captures images and builds DOCX
- PyVista off-screen rendering (`pv.OFF_SCREEN = True`) for 3D contour screenshots
- Captures ~30 images: mesh views, deformed mode shapes, modal stress, 1-sigma displacement/stress contours from multiple angles (iso, front, side, top)
- 1-sigma PSD results plotted by assigning manually-computed arrays to `grid.point_data` on `mapdl.mesh.grid`
- 4 matplotlib charts: frequencies, PSD input, modal contribution, composite layup
- Generates professional DOCX via `python-docx` with tables, figure captions, blue heading theme
- Output: `report_output/PSD_Analysis_Report.docx` + 30 PNG images
- Typical run: ~15 seconds total

### MCP Server (`ansys_mcp_server.py`)
- Uses `fastmcp` framework, communicates via stdio
- Maintains a global `mapdl` session object (singleton pattern via `launch_mapdl()`)
- Tools organized by workflow phase: session -> geometry -> materials -> elements -> meshing -> BCs -> solving -> post-processing
- `import_geometry()` uses ac4 converter for Parasolid/STEP files
- Arbitrary APDL commands via `run_apdl_command()` / `run_apdl_commands()`
- All units: SI (metres, Pascals, kg/m3, Hz)

### Legacy Script (`composite_random_vibration.py`)
- Uses ANF commands fallback for geometry (broken topology — produces volume but mesh fails)
- SOLID186 KEYOPT(3)=1 layered solids (requires hex mesh, incompatible with wrench geometry)
- SECTYPE SOLID syntax broken in v25.2 Student

## Key Conventions

- Material properties are orthotropic (EX, EY, EZ, GXY, GXZ, GYZ, PRXY, PRXZ, PRYZ, DENS)
- Composite sections defined via SECTYPE/SECDATA with explicit ply-by-ply specification
- The MCP server returns structured dicts/strings; errors are caught and returned as descriptive messages
- APDL commands embedded in PyMAPDL use `mapdl.run()` for raw command strings

## ANSYS v25.2 Student Edition Quirks (CRITICAL)

These are confirmed issues specific to Student 2025 R2 (v25.2):

| Issue | Symptom | Fix |
|-------|---------|-----|
| **Parasolid import** | `~PARAIN` fails via gRPC | Use `ac4para.exe` with `P_SCHEMA` env var pointing to `v252/commonfiles/CAD/Siemens/Parasolid36.1.227/winx64/schema` |
| **ANF loading** | `mapdl.run("/INPUT,file,anf")` silently returns nothing | Use `mapdl.input(full_path)` instead |
| **SECTYPE SOLID** | `SECTYPE,1,SOLID,,Name` -> "not a valid solid section subtype" | Use `SECTYPE,1,SHELL,,Name` (works with SOLID186 layered) |
| **SOLID186 layered mesh** | KEYOPT(3)=1 requires mapped hex mesh, fails on complex geometry | Use SOLID187 (10-node tet) with orthotropic properties instead |
| **PSD 1-sigma results** | SET(2,1) stores placeholder UY=1.0 for all nodes; "VERIFICATION RUN ONLY" | Compute 1-sigma manually from modal data using SRSS |
| **numpy v2.x** | `np.trapz` removed | Use `np.trapezoid` |
| **Unicode print** | Windows cp1252 can't encode arrow/special chars | Use ASCII alternatives in print statements |
| **ANF commands.json** | Produces volume entity (1 vol, 18 areas) but topology is broken — mesh fails with "no exterior faces" | Use ac4 converter instead |
| **IGES import** | `heavyDutyWrench.iges` -> "Poorly defined area" / "crossed lines" | Use Parasolid via ac4 instead |

## Common Issues

- **"Cannot find Ansys installation"**: Set `exec_file` path manually in `launch_mapdl()` call
- **Mesh exceeds student license**: Increase `ELEMENT_SIZE` constant (e.g., `0.005` for 5mm elements)
- **Stale MAPDL process**: Kill with `taskkill //F //IM ANSYS252.exe` (use `//` in Git Bash) or use `close_mapdl()` MCP tool
- **`python` not found**: Use full path `C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe`
- **gRPC connection timeout**: Kill all stale ANSYS processes first, wait 5 seconds, then retry
