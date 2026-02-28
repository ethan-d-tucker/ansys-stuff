# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

ANSYS composite random vibration (PSD) analysis project for a heavy-duty wrench with sandwich composite construction. Two main components:

1. **`ansys_mcp_server.py`** — FastMCP server exposing 30+ ANSYS MAPDL tools for interactive Claude-driven FEA workflows
2. **`composite_random_vibration.py`** — Standalone end-to-end analysis script (geometry → materials → mesh → modal → PSD → results)

Supporting files: `generate_plots.py` (matplotlib visualization), `psd_curve.csv` (input spectrum), `heavyDutyWrench.iges` (CAD geometry), `anf_commands.json` (ANF geometry fallback).

## Environment Setup

```bash
python -m venv ansys_env
ansys_env\Scripts\activate
pip install ansys-mapdl-core fastmcp matplotlib pyvista numpy
```

Requires **ANSYS Student 2025/2025R2** installed locally. Student license limits mesh to ~128k nodes/elements.

## Running

```bash
# Full standalone analysis
python composite_random_vibration.py

# Generate result plots from hardcoded data
python generate_plots.py

# MCP server is auto-launched by Claude Code via .mcp.json
```

## Architecture

### MCP Server (`ansys_mcp_server.py`)
- Uses `fastmcp` framework, communicates via stdio
- Maintains a global `mapdl` session object (singleton pattern via `launch_mapdl()`)
- Tools are organized by workflow phase: session → geometry → materials → elements → meshing → BCs → solving → post-processing
- Arbitrary APDL commands available via `run_apdl_command()` / `run_apdl_commands()` escape hatches
- All units: SI (metres, Pascals, kg/m³, Hz)

### Analysis Script (`composite_random_vibration.py`)
- All configurable parameters are constants at the top of the file (geometry path, material properties, ply thicknesses, solver settings)
- Composite layup: symmetric sandwich — 5 carbon/epoxy plies + honeycomb core + 5 carbon/epoxy plies
- Element type: SOLID186 with KEYOPT(3)=1 for layered solids
- Fixed BC applied at the handle end (auto-detected as the axis extremum)
- Modal solve (Block Lanczos, 20 modes, 0–3000 Hz) feeds into PSD spectrum analysis (Y-axis base excitation, 2% damping)

## Key Conventions

- Material properties are orthotropic (EX, EY, EZ, GXY, GXZ, GYZ, PRXY, PRXZ, PRYZ, DENS)
- Composite sections defined via SECTYPE/SECDATA with explicit ply-by-ply specification
- The MCP server returns structured dicts/strings; errors are caught and returned as descriptive messages rather than raising exceptions
- APDL commands embedded in PyMAPDL use `mapdl.run()` for raw command strings

## Common Issues

- **"Cannot find Ansys installation"**: Set `exec_file` path manually in `launch_mapdl()` call
- **Mesh exceeds student license**: Increase `ELEMENT_SIZE` constant (e.g., `0.005` for 5mm elements)
- **Stale MAPDL process**: Kill with `taskkill /F /IM ANSYS252.exe` or use `close_mapdl()` MCP tool before relaunching
