# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ANSYS composite random vibration (PSD) analysis project for a heavy-duty wrench with sandwich composite construction. Four main components:

1. **`ansys_mcp_server.py`** — FastMCP server exposing 30+ ANSYS MAPDL tools for interactive Claude-driven FEA workflows
2. **`composite_random_vibration.py`** — Standalone end-to-end analysis script (geometry -> materials -> mesh -> modal -> PSD -> results). Uses ANF fallback for geometry import (has known issues, see below)
3. **`run_simulation.py`** — Working end-to-end simulation using all debugged/validated methods. **Use this script for actual runs.**
4. **`run_and_report.py`** — Full simulation + DOCX report generator. Runs the PSD analysis, captures ~30 PyVista 3D contour screenshots (mesh, mode shapes, 1-sigma stress/displacement from multiple angles), generates matplotlib charts, and builds a professional Word report. **Use this for reports.**

Supporting files: `generate_plots.py` (matplotlib visualization), `generate_report.py` (HTML report generator), `psd_curve.csv` (input spectrum), `WrenchParasolid.x_t` (primary CAD geometry), `heavyDutyWrench.iges` (IGES fallback), `anf_commands.json` (ANF geometry fallback — has broken topology, avoid).

## Environment

- **Python**: `C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe` (system install, no venv)
- **ANSYS**: Student 2025 R2 (v25.2) at `C:\Program Files\ANSYS Inc\ANSYS Student\v252`
- **Packages**: ansys-mapdl-core, fastmcp, matplotlib, pyvista, numpy (v2.x), python-docx
- **Shell**: Git Bash on Windows — use `//` for flag prefixes in taskkill (e.g., `taskkill //F //IM ANSYS252.exe`)
- Student license limits mesh to ~128k nodes/elements

## Running

```bash
# Working simulation (recommended)
"C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe" run_simulation.py

# Full simulation + DOCX report with 3D contour plots (~15s)
"C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe" run_and_report.py
# Output: report_output/PSD_Analysis_Report.docx (30 images, ~15 pages)

# Generate result plots from hardcoded data
"C:/Users/Ethan/AppData/Local/Programs/Python/Python311/python.exe" generate_plots.py

# MCP server is auto-launched by Claude Code via .mcp.json
```

## Architecture

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
