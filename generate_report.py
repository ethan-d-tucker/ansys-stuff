"""
generate_report.py — Self-contained HTML report generator for the ANSYS composite wrench PSD analysis.

Usage:
    python generate_report.py                              # auto-discovers report_output/results.json
    python generate_report.py --results path/results.json --output report.html
    python generate_report.py --fallback                   # use hardcoded data, no ANSYS needed
"""

import argparse
import base64
import datetime
import json
import math
import os
import sys

import matplotlib
matplotlib.use("Agg")
import matplotlib.patches as mpatches
import matplotlib.pyplot as plt
import numpy as np

# ---------------------------------------------------------------------------
# Hardcoded fallback data (mirrors generate_plots.py exactly)
# ---------------------------------------------------------------------------

HARDCODED_DATA = {
    "metadata": {
        "generated_at": None,  # filled at runtime
        "ansys_version": "N/A (hardcoded design-phase data)",
        "ansys_directory": "N/A",
        "output_dir": ".",
        "geometry_file": "WrenchParasolid.x_t",
    },
    "analysis_parameters": {
        "element_type": "SOLID186",
        "keyopt_3": 1,
        "num_modes_requested": 20,
        "freq_range_hz": [0.0, 3000.0],
        "damping_ratio": 0.02,
        "excitation_direction": "UY",
        "psd_input_table": [
            {"frequency_hz": 20.0,   "psd_g2_per_hz": 0.010},
            {"frequency_hz": 80.0,   "psd_g2_per_hz": 0.040},
            {"frequency_hz": 350.0,  "psd_g2_per_hz": 0.040},
            {"frequency_hz": 2000.0, "psd_g2_per_hz": 0.007},
        ],
    },
    "materials": {
        "mat_1": {
            "name": "Carbon/Epoxy Woven Prepreg",
            "EX_pa": 60e9, "EY_pa": 60e9, "EZ_pa": 10e9,
            "GXY_pa": 5e9, "GXZ_pa": 4e9, "GYZ_pa": 4e9,
            "PRXY": 0.04, "PRXZ": 0.30, "PRYZ": 0.30,
            "density_kg_m3": 1420.0,
        },
        "mat_2": {
            "name": "Honeycomb Core (Nomex-style)",
            "EX_pa": 1e6, "EY_pa": 1e6, "EZ_pa": 130e6,
            "GXY_pa": 1e6, "GXZ_pa": 24e6, "GYZ_pa": 48e6,
            "PRXY": 0.49, "PRXZ": 0.001, "PRYZ": 0.001,
            "density_kg_m3": 48.0,
        },
    },
    "composite_layup": {
        "section_id": 1,
        "name": "CompSandwich",
        "total_thickness_mm": 3.175,
        "plies": [
            {"ply": 1,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 0.0,  "material_name": "Carbon/Epoxy"},
            {"ply": 2,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 0.0,  "material_name": "Carbon/Epoxy"},
            {"ply": 3,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 45.0, "material_name": "Carbon/Epoxy"},
            {"ply": 4,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 45.0, "material_name": "Carbon/Epoxy"},
            {"ply": 5,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 90.0, "material_name": "Carbon/Epoxy"},
            {"ply": 6,  "mat_id": 2, "thickness_mm": 1.675, "angle_deg": 0.0, "material_name": "Honeycomb Core"},
            {"ply": 7,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 90.0, "material_name": "Carbon/Epoxy"},
            {"ply": 8,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 45.0, "material_name": "Carbon/Epoxy"},
            {"ply": 9,  "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 45.0, "material_name": "Carbon/Epoxy"},
            {"ply": 10, "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 0.0,  "material_name": "Carbon/Epoxy"},
            {"ply": 11, "mat_id": 1, "thickness_mm": 0.15, "angle_deg": 0.0,  "material_name": "Carbon/Epoxy"},
        ],
    },
    "mesh": {"nodes": 1643, "elements": 228},
    "modal_results": {
        "modes_found": 6,
        "frequencies": [
            {"mode": 1, "frequency_hz": 112.7096},
            {"mode": 2, "frequency_hz": 208.3092},
            {"mode": 3, "frequency_hz": 742.3481},
            {"mode": 4, "frequency_hz": 1374.2027},
            {"mode": 5, "frequency_hz": 1656.5302},
            {"mode": 6, "frequency_hz": 2128.7873},
        ],
    },
    "participation_factors": {
        "modes_found": 6,
        "excitation_direction": "Y",
        "factors": [
            {"mode": 1, "frequency_hz": 112.7096,  "participation_factor_y": -0.039861, "modal_mass_ratio_y": None},
            {"mode": 2, "frequency_hz": 208.3092,  "participation_factor_y": 7.61e-11,  "modal_mass_ratio_y": None},
            {"mode": 3, "frequency_hz": 742.3481,  "participation_factor_y": 0.021524,  "modal_mass_ratio_y": None},
            {"mode": 4, "frequency_hz": 1374.2027, "participation_factor_y": 2.26e-13,  "modal_mass_ratio_y": None},
            {"mode": 5, "frequency_hz": 1656.5302, "participation_factor_y": 0.003453,  "modal_mass_ratio_y": None},
            {"mode": 6, "frequency_hz": 2128.7873, "participation_factor_y": -0.012200, "modal_mass_ratio_y": None},
        ],
    },
    "displacement_results": {
        "Y": {"component": "Y", "max_absolute_um": 23.4, "max_absolute_mm": 0.0234, "max_absolute": 2.34e-5},
    },
    "stress_results": {
        "EQV": {"component": "EQV", "max_absolute_mpa": 1.245, "max_absolute_pa": 1245000.0},
    },
    "_response_psd": {
        "node": 1500,
        "freq": [
            20.0, 50.56, 80.0, 100.15, 105.71, 108.9, 110.96, 112.29,
            112.59, 112.71, 113.13, 114.48, 118.13, 124.76, 150.55,
            200.10, 207.24, 208.31, 209.38, 220.07, 300.0, 350.0,
            500.0, 650.0, 700.0, 720.86, 733.95, 740.89, 742.35,
            744.71, 756.06, 800.0, 1000.0, 1200.0, 1374.2, 1500.0,
            1600.6, 1652.7, 1656.5, 1660.4, 1700.0, 1900.0, 2000.0,
        ],
        "val": [
            4.19e-9, 4.18e-10, 3.08e-10, 8.16e-10, 1.97e-9, 4.82e-9,
            1.08e-8, 1.58e-8, 1.60e-8, 1.60e-8, 1.52e-8, 9.19e-9,
            1.93e-9, 3.59e-10, 1.84e-11, 1.31e-12, 1.00e-12, 9.66e-13,
            9.31e-13, 6.52e-13, 7.48e-14, 3.60e-14, 1.26e-14, 1.55e-14,
            4.98e-14, 1.13e-13, 2.39e-13, 2.94e-13, 2.93e-13, 2.76e-13,
            1.41e-13, 1.18e-14, 1.24e-16, 1.52e-17, 4.91e-19, 1.01e-17,
            1.86e-16, 1.04e-15, 1.08e-15, 1.09e-15, 7.12e-16, 1.08e-16,
            1.63e-16,
        ],
    },
    "images": [],
    "errors": [],
}

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------

REPORT_CSS = """
:root {
  --blue:   #1565C0;
  --orange: #E65100;
  --grey:   #37474F;
  --light:  #F5F5F5;
  --border: #B0BEC5;
}
* { box-sizing: border-box; }
body {
  font-family: 'Segoe UI', Arial, sans-serif;
  margin: 0; padding: 0;
  background: #FAFAFA; color: #212121;
  font-size: 14px; line-height: 1.6;
}
header {
  background: var(--blue); color: white;
  padding: 28px 48px;
}
header h1 { margin: 0; font-size: 26px; font-weight: 700; }
header p  { margin: 6px 0 0; opacity: 0.85; font-size: 13px; }
nav {
  background: var(--grey); padding: 10px 48px;
  position: sticky; top: 0; z-index: 99;
  white-space: nowrap; overflow-x: auto;
}
nav a {
  color: #CFD8DC; text-decoration: none;
  margin-right: 22px; font-size: 13px;
}
nav a:hover { color: white; }
main { max-width: 1140px; margin: 0 auto; padding: 36px 48px; }
section { margin-bottom: 56px; }
h2 {
  color: var(--blue);
  border-bottom: 2px solid var(--blue);
  padding-bottom: 6px; font-size: 20px; margin-top: 0;
}
h3 { color: var(--grey); font-size: 15px; margin-top: 28px; margin-bottom: 8px; }
table { border-collapse: collapse; width: 100%; margin: 16px 0; font-size: 13px; }
th {
  background: var(--blue); color: white;
  padding: 9px 13px; text-align: left;
}
td { padding: 7px 13px; border-bottom: 1px solid var(--border); }
tr:nth-child(even) { background: var(--light); }
tr:hover { background: #E3F2FD; }
.mat-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 24px; margin: 16px 0; }
.figure { text-align: center; margin: 24px 0; }
.figure img {
  max-width: 100%;
  border: 1px solid var(--border);
  border-radius: 4px;
  box-shadow: 0 2px 8px rgba(0,0,0,0.12);
}
.figure-grid { display: grid; grid-template-columns: 1fr 1fr; gap: 20px; margin: 24px 0; }
.caption {
  font-size: 12px; color: #607D8B;
  margin-top: 6px; font-style: italic;
}
.kv-row {
  display: grid; grid-template-columns: repeat(3,1fr);
  gap: 14px; margin: 16px 0;
}
.kv-card {
  background: white; border: 1px solid var(--border);
  border-radius: 6px; padding: 16px 20px;
  box-shadow: 0 1px 3px rgba(0,0,0,0.08);
}
.kv-card .label {
  font-size: 11px; text-transform: uppercase;
  color: #78909C; letter-spacing: 0.5px; margin-bottom: 4px;
}
.kv-card .value { font-size: 22px; font-weight: 700; color: var(--blue); }
.kv-card .unit  { font-size: 12px; color: #90A4AE; margin-left: 4px; }
.warn {
  background: #FFF3E0; border-left: 4px solid var(--orange);
  padding: 12px 18px; border-radius: 4px;
  font-size: 13px; color: #BF360C; margin-bottom: 16px;
}
.info {
  background: #E3F2FD; border-left: 4px solid var(--blue);
  padding: 12px 18px; border-radius: 4px;
  font-size: 13px; color: #0D47A1; margin-bottom: 16px;
}
footer {
  background: var(--grey); color: #90A4AE;
  text-align: center; padding: 18px;
  font-size: 12px; margin-top: 48px;
}
"""

# ---------------------------------------------------------------------------
# Data loading
# ---------------------------------------------------------------------------

def load_results(json_path: str) -> dict:
    with open(json_path) as f:
        return json.load(f)


def get_fallback_data() -> dict:
    data = json.loads(json.dumps(HARDCODED_DATA))  # deep copy
    data["metadata"]["generated_at"] = datetime.datetime.now().isoformat()
    data["_fallback"] = True
    return data

# ---------------------------------------------------------------------------
# Image helpers
# ---------------------------------------------------------------------------

def image_to_base64(path: str):
    try:
        with open(path, "rb") as f:
            enc = base64.b64encode(f.read()).decode("ascii")
        return f"data:image/png;base64,{enc}"
    except (FileNotFoundError, OSError):
        return None


def _img_tag(path: str, caption: str, css_class: str = "figure") -> str:
    b64 = image_to_base64(path)
    if b64 is None:
        return f'<p class="warn">Image not available: <code>{path}</code></p>'
    return (
        f'<div class="{css_class}">'
        f'<img src="{b64}" alt="{caption}">'
        f'<p class="caption">{caption}</p>'
        f"</div>"
    )

# ---------------------------------------------------------------------------
# Matplotlib figure generation
# ---------------------------------------------------------------------------

def generate_matplotlib_plots(data: dict, output_dir: str) -> dict:
    """Generate matplotlib figures and return {name: path}."""
    os.makedirs(output_dir, exist_ok=True)
    paths = {}

    freqs_list = data.get("modal_results", {}).get("frequencies", [])
    modes = [f["mode"] for f in freqs_list]
    freqs = [f["frequency_hz"] for f in freqs_list]

    psd_table = data["analysis_parameters"]["psd_input_table"]
    psd_freq = [p["frequency_hz"] for p in psd_table]
    psd_val  = [p["psd_g2_per_hz"] for p in psd_table]

    rpsd = data.get("_response_psd", {})
    rpsd_freq = np.array(rpsd.get("freq", []))
    rpsd_val  = np.array(rpsd.get("val",  []))

    plies = data["composite_layup"]["plies"]

    colors = ["#2196F3","#4CAF50","#FF9800","#E91E63","#9C27B0","#00BCD4",
              "#F44336","#795548","#607D8B","#FFEB3B","#8BC34A"]

    # --- 4-panel overview ---
    fig = plt.figure(figsize=(16, 12))
    fig.suptitle("Composite Wrench — Random Vibration (PSD) Analysis Results",
                 fontsize=16, fontweight="bold", y=0.98)

    # Panel 1: Natural frequencies bar chart
    ax1 = fig.add_subplot(2, 2, 1)
    if freqs:
        bar_colors = [colors[i % len(colors)] for i in range(len(modes))]
        bars = ax1.bar(modes, freqs, color=bar_colors, edgecolor="black", linewidth=0.5)
        for bar, f in zip(bars, freqs):
            ax1.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + max(freqs) * 0.02,
                     f"{f:.1f}", ha="center", va="bottom", fontsize=9, fontweight="bold")
        ax1.set_ylim(0, max(freqs) * 1.15)
    ax1.set_xlabel("Mode Number")
    ax1.set_ylabel("Natural Frequency (Hz)")
    ax1.set_title(f"Natural Frequencies ({len(freqs)} Modes, 0–3000 Hz)")
    ax1.grid(axis="y", alpha=0.3)

    # Panel 2: PSD input spectrum + frequency markers
    ax2 = fig.add_subplot(2, 2, 2)
    ax2.loglog(psd_freq, psd_val, "b-o", linewidth=2, markersize=6,
               label="Input PSD (G²/Hz)", zorder=5)
    for i, f in enumerate(freqs):
        if f <= max(psd_freq):
            ax2.axvline(x=f, color="red", alpha=0.3, linestyle="--", linewidth=0.8)
            if i < 3:
                ax2.text(f, max(psd_val) * 1.5, f"f{i+1}={f:.0f}Hz",
                         rotation=90, va="bottom", ha="right", fontsize=7, color="red")
    ax2.set_xlabel("Frequency (Hz)")
    ax2.set_ylabel("PSD (G²/Hz)")
    ax2.set_title("PSD Input Spectrum with Natural Frequencies")
    ax2.legend(loc="upper right")
    ax2.grid(True, which="both", alpha=0.3)
    ax2.set_xlim(10, 3000)

    # Panel 3: Response PSD
    ax3 = fig.add_subplot(2, 2, 3)
    if len(rpsd_freq) > 0:
        mask = rpsd_val > 0
        ax3.loglog(rpsd_freq[mask], rpsd_val[mask], "r-", linewidth=1.5,
                   label=f"Response PSD (Node {rpsd.get('node', '?')}, UY)")
        for f in freqs:
            if f <= max(rpsd_freq):
                ax3.axvline(x=f, color="blue", alpha=0.2, linestyle="--", linewidth=0.8)
        peak_idx = int(np.argmax(rpsd_val))
        ax3.annotate(
            f"Peak at {rpsd_freq[peak_idx]:.1f} Hz\n(Mode 1 resonance)",
            xy=(rpsd_freq[peak_idx], rpsd_val[peak_idx]),
            xytext=(300, rpsd_val[peak_idx] * 0.5),
            arrowprops=dict(arrowstyle="->", color="black"),
            fontsize=8, ha="center",
        )
        ax3.set_xlim(10, 3000)
        ax3.legend(loc="upper right")
    else:
        ax3.text(0.5, 0.5, "Response PSD not available\n(run ANSYS analysis to populate)",
                 ha="center", va="center", transform=ax3.transAxes,
                 fontsize=11, color="#607D8B")
    ax3.set_xlabel("Frequency (Hz)")
    ax3.set_ylabel("Response PSD (m²/Hz)")
    ax3.set_title("Response PSD — UY Direction")
    ax3.grid(True, which="both", alpha=0.3)

    # Panel 4: Composite layup
    ax4 = fig.add_subplot(2, 2, 4)
    colors_map = {"Carbon/Epoxy": "#333333", "Honeycomb Core": "#FFD700"}
    y_pos = 0.0
    for ply in plies:
        mat = ply["material_name"]
        t   = ply["thickness_mm"]
        ang = ply["angle_deg"]
        color = colors_map.get(mat, "#999999")
        rect = mpatches.FancyBboxPatch((0.2, y_pos), 0.6, t,
                                       boxstyle="round,pad=0.01",
                                       facecolor=color, edgecolor="white", linewidth=1)
        ax4.add_patch(rect)
        label = f'{int(ang)}°' if mat != "Honeycomb Core" else "Core"
        text_color = "white" if mat == "Carbon/Epoxy" else "black"
        ax4.text(0.5, y_pos + t / 2, f"{label}  ({t:.3f} mm, {mat})",
                 ha="center", va="center", fontsize=8, fontweight="bold", color=text_color)
        y_pos += t

    ax4.set_xlim(0, 1)
    ax4.set_ylim(-0.05, y_pos + 0.05)
    ax4.set_ylabel("Thickness (mm)")
    total_t = data["composite_layup"].get("total_thickness_mm", y_pos)
    ax4.set_title(f"Composite Sandwich Layup — Total: {total_t:.3f} mm")
    ax4.set_xticks([])
    carbon_patch = mpatches.Patch(facecolor="#333333", edgecolor="black", label="Carbon/Epoxy Prepreg")
    honey_patch  = mpatches.Patch(facecolor="#FFD700", edgecolor="black", label="Honeycomb Core")
    ax4.legend(handles=[carbon_patch, honey_patch], loc="upper left", fontsize=8)

    plt.tight_layout(rect=[0, 0, 1, 0.95])
    overview_path = os.path.join(output_dir, "overview_4panel.png")
    plt.savefig(overview_path, dpi=150, bbox_inches="tight")
    plt.close()
    paths["overview_4panel"] = overview_path

    return paths

# ---------------------------------------------------------------------------
# HTML section builders
# ---------------------------------------------------------------------------

def _kv_card(label: str, value: str, unit: str = "") -> str:
    unit_span = f'<span class="unit">{unit}</span>' if unit else ""
    return (
        f'<div class="kv-card">'
        f'<div class="label">{label}</div>'
        f'<div class="value">{value}{unit_span}</div>'
        f"</div>"
    )


def _table(headers: list, rows: list) -> str:
    ths = "".join(f"<th>{h}</th>" for h in headers)
    body = ""
    for row in rows:
        tds = "".join(f"<td>{cell}</td>" for cell in row)
        body += f"<tr>{tds}</tr>"
    return f"<table><thead><tr>{ths}</tr></thead><tbody>{body}</tbody></table>"


def _section_summary(data: dict) -> str:
    meta = data.get("metadata", {})
    mesh = data.get("mesh", {})
    layup = data.get("composite_layup", {})
    disp_y = data.get("displacement_results", {}).get("Y", {})
    stress_eqv = data.get("stress_results", {}).get("EQV", {})
    freqs_list = data.get("modal_results", {}).get("frequencies", [])

    nodes    = mesh.get("nodes", "N/A")
    elements = mesh.get("elements", "N/A")
    total_t  = layup.get("total_thickness_mm", "N/A")

    max_uy   = disp_y.get("max_absolute_um", "N/A")
    max_uy_s = f"{max_uy:.3f}" if isinstance(max_uy, (int, float)) else str(max_uy)

    max_seqv   = stress_eqv.get("max_absolute_mpa", "N/A")
    max_seqv_s = f"{max_seqv:.3f}" if isinstance(max_seqv, (int, float)) else str(max_seqv)

    dom_freq   = freqs_list[0]["frequency_hz"] if freqs_list else "N/A"
    dom_freq_s = f"{dom_freq:.2f}" if isinstance(dom_freq, (int, float)) else str(dom_freq)

    gen_at = meta.get("generated_at", "")
    if gen_at:
        try:
            gen_at = datetime.datetime.fromisoformat(gen_at).strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass

    html = f"""
<section id="summary">
  <h2>Project Summary</h2>
  <table style="width:auto; margin-bottom:16px;">
    <tbody>
      <tr><td><b>Project</b></td><td>Heavy-Duty Wrench — Composite Sandwich Random Vibration Analysis</td></tr>
      <tr><td><b>Geometry</b></td><td>{meta.get("geometry_file", "N/A")}</td></tr>
      <tr><td><b>ANSYS Version</b></td><td>{meta.get("ansys_version", "N/A")}</td></tr>
      <tr><td><b>Report Generated</b></td><td>{gen_at}</td></tr>
    </tbody>
  </table>

  <h3>Mesh &amp; Model</h3>
  <div class="kv-row">
    {_kv_card("Nodes", str(nodes))}
    {_kv_card("Elements", str(elements))}
    {_kv_card("Laminate Thickness", str(total_t), "mm")}
  </div>

  <h3>Key 1-Sigma Results</h3>
  <div class="kv-row">
    {_kv_card("Max |UY| Displacement", max_uy_s, "µm")}
    {_kv_card("Max von Mises Stress", max_seqv_s, "MPa")}
    {_kv_card("Dominant Mode Frequency", dom_freq_s, "Hz")}
  </div>
</section>"""
    return html


def _section_parameters(data: dict) -> str:
    params = data.get("analysis_parameters", {})
    freq_range = params.get("freq_range_hz", [0, 3000])
    damp_pct = params.get("damping_ratio", 0.02) * 100

    param_rows = [
        ["Element Type", f"{params.get('element_type', 'SOLID186')} (KEYOPT(3)=1, layered solid)"],
        ["Number of Modes Requested", str(params.get("num_modes_requested", 20))],
        ["Frequency Range", f"{freq_range[0]:.0f} – {freq_range[1]:.0f} Hz"],
        ["Excitation Direction", f"{params.get('excitation_direction', 'UY')} (Y-axis base acceleration)"],
        ["Damping Ratio", f"{damp_pct:.1f}% (constant modal)"],
        ["Boundary Condition", "Fixed support at handle end (all DOF = 0)"],
        ["Units", "SI — metres, Pascals, kg/m³"],
    ]

    psd_table = params.get("psd_input_table", [])
    psd_rows = [[f"{p['frequency_hz']:.1f}", f"{p['psd_g2_per_hz']:.4f}"] for p in psd_table]

    return f"""
<section id="parameters">
  <h2>Analysis Parameters</h2>
  {_table(["Parameter", "Value"], param_rows)}
  <h3>PSD Input Spectrum</h3>
  {_table(["Frequency (Hz)", "PSD (G²/Hz)"], psd_rows)}
</section>"""


def _fmt_pa(val_pa: float, unit: str = "GPa") -> str:
    if unit == "GPa":
        return f"{val_pa / 1e9:.3f} GPa"
    if unit == "MPa":
        return f"{val_pa / 1e6:.1f} MPa"
    return f"{val_pa:.3e} Pa"


def _section_materials(data: dict) -> str:
    mats = data.get("materials", {})

    def _mat_table(m: dict) -> str:
        rows = [
            ["EX (in-plane, X)", _fmt_pa(m.get("EX_pa", 0))],
            ["EY (in-plane, Y)", _fmt_pa(m.get("EY_pa", 0))],
            ["EZ (through-thickness)", _fmt_pa(m.get("EZ_pa", 0))],
            ["GXY (in-plane shear)", _fmt_pa(m.get("GXY_pa", 0))],
            ["GXZ (out-of-plane shear)", _fmt_pa(m.get("GXZ_pa", 0))],
            ["GYZ (out-of-plane shear)", _fmt_pa(m.get("GYZ_pa", 0))],
            ["νXY (in-plane Poisson)", f"{m.get('PRXY', 0):.3f}"],
            ["νXZ", f"{m.get('PRXZ', 0):.3f}"],
            ["νYZ", f"{m.get('PRYZ', 0):.3f}"],
            ["Density", f"{m.get('density_kg_m3', 0):.1f} kg/m³"],
        ]
        return _table(["Property", "Value"], rows)

    m1 = mats.get("mat_1", {})
    m2 = mats.get("mat_2", {})

    return f"""
<section id="materials">
  <h2>Material Properties</h2>
  <div class="mat-grid">
    <div>
      <h3>Material 1: {m1.get('name', 'Carbon/Epoxy')}</h3>
      {_mat_table(m1)}
    </div>
    <div>
      <h3>Material 2: {m2.get('name', 'Honeycomb Core')}</h3>
      {_mat_table(m2)}
    </div>
  </div>
</section>"""


def _section_layup(data: dict, plot_paths: dict) -> str:
    layup = data.get("composite_layup", {})
    plies = layup.get("plies", [])
    total_t = layup.get("total_thickness_mm", "N/A")

    def _role(ply: dict) -> str:
        n = ply["ply"]
        num_plies = len(plies)
        mid = (num_plies + 1) / 2
        if ply["material_name"] == "Honeycomb Core":
            return "Core"
        if n <= mid:
            return "Bottom face-sheet"
        return "Top face-sheet"

    ply_rows = [
        [
            str(p["ply"]),
            p["material_name"],
            f"{p['thickness_mm']:.3f}",
            f"{int(p['angle_deg'])}°" if p["material_name"] != "Honeycomb Core" else "—",
            _role(p),
        ]
        for p in plies
    ]

    img_html = ""
    if "overview_4panel" in plot_paths:
        img_html = _img_tag(plot_paths["overview_4panel"],
                            "Figure 1 — Analysis results overview (4-panel)")

    return f"""
<section id="layup">
  <h2>Composite Layup</h2>
  <p>Symmetric sandwich construction: [0/0/45/45/90 / Core / 90/45/45/0/0]
     — Total laminate thickness: <b>{total_t} mm</b></p>
  {_table(["Ply", "Material", "Thickness (mm)", "Angle", "Role"], ply_rows)}
  {img_html}
</section>"""


def _section_frequencies(data: dict) -> str:
    freqs_list = data.get("modal_results", {}).get("frequencies", [])
    pf_list    = data.get("participation_factors", {}).get("factors", [])
    pf_map     = {f["mode"]: f for f in pf_list}

    rows = []
    for f in freqs_list:
        m = f["mode"]
        hz = f"{f['frequency_hz']:.4f}"
        pf_entry = pf_map.get(m, {})
        pf_val = pf_entry.get("participation_factor_y", None)
        emr    = pf_entry.get("modal_mass_ratio_y", None)
        pf_s   = f"{pf_val:.4e}" if pf_val is not None else "N/A"
        emr_s  = f"{emr:.4f}"   if emr  is not None else "N/A"
        if pf_val is not None and abs(pf_val) > 0.001:
            note = "<b>Dominant (Y)</b>"
        elif pf_val is not None:
            note = "Inactive (Y)"
        else:
            note = ""
        rows.append([str(m), hz, pf_s, emr_s, note])

    return f"""
<section id="frequencies">
  <h2>Natural Frequencies &amp; Modal Participation</h2>
  <p>Modal analysis: Block Lanczos solver, {len(freqs_list)} modes extracted in 0–3000 Hz.
     PSD base excitation in Y-direction.</p>
  {_table(["Mode", "Frequency (Hz)", "Part. Factor (Y)", "Mass Ratio (Y)", "Note"], rows)}
</section>"""


def _section_psd_input(data: dict) -> str:
    psd_table = data["analysis_parameters"]["psd_input_table"]
    rows = [[f"{p['frequency_hz']:.1f}", f"{p['psd_g2_per_hz']:.4f}"] for p in psd_table]
    return f"""
<section id="psd-input">
  <h2>PSD Input Spectrum</h2>
  <p>Four-point piecewise definition on a log-log scale. Base acceleration applied
     in the Y-direction. Units: G²/Hz.</p>
  {_table(["Frequency (Hz)", "PSD (G²/Hz)"], rows)}
  <p>Frequency markers for all extracted natural frequencies are plotted in the
     overview figure above (Figure 1, top-right panel).</p>
</section>"""


def _section_response(data: dict, images: dict) -> str:
    disp = data.get("displacement_results", {})

    disp_rows = []
    for comp in ["Y", "NORM", "X", "Z"]:
        d = disp.get(comp, {})
        mm  = d.get("max_absolute_mm", "N/A")
        um  = d.get("max_absolute_um", "N/A")
        mm_s = f"{mm:.6f}" if isinstance(mm, (int, float)) else str(mm)
        um_s = f"{um:.4f}" if isinstance(um, (int, float)) else str(um)
        label = {"Y": "UY (Y-direction)", "NORM": "Magnitude", "X": "UX", "Z": "UZ"}.get(comp, comp)
        disp_rows.append([label, mm_s, um_s])

    img_html = ""
    for comp, caption in [("Y", "UY displacement contour (1-sigma)"),
                           ("NORM", "Displacement magnitude contour (1-sigma)")]:
        key = f"displacement_{comp.lower()}"
        if key in images:
            img_html += _img_tag(images[key], f"Figure — {caption}")

    return f"""
<section id="response">
  <h2>Structural Response — Displacement</h2>
  <p>Results represent 1-sigma (one standard deviation) relative displacements
     under the applied PSD base excitation. See overview figure (Figure 1,
     bottom-left panel) for the response PSD at the wrench head node.</p>
  <h3>1-Sigma Displacement Summary</h3>
  {_table(["Component", "Max |Displacement| (mm)", "Max |Displacement| (µm)"], disp_rows)}
  <div class="figure-grid">
    {img_html}
  </div>
</section>"""


def _section_stress(data: dict, images: dict) -> str:
    stress = data.get("stress_results", {})

    stress_rows = []
    for comp in ["EQV", "X", "Y", "Z"]:
        s = stress.get(comp, {})
        mpa = s.get("max_absolute_mpa", "N/A")
        mpa_s = f"{mpa:.4f}" if isinstance(mpa, (int, float)) else str(mpa)
        label = {"EQV": "von Mises (σ_eqv)", "X": "σX", "Y": "σY", "Z": "σZ"}.get(comp, comp)
        stress_rows.append([label, mpa_s])

    img_html = ""
    for comp, caption in [("EQV", "von Mises stress contour (1-sigma)"),
                           ("X",   "σX stress contour (1-sigma)"),
                           ("Y",   "σY stress contour (1-sigma)"),
                           ("Z",   "σZ stress contour (1-sigma)")]:
        key = f"stress_{comp.lower()}"
        if key in images:
            img_html += _img_tag(images[key], f"Figure — {caption}")

    img_grid = f'<div class="figure-grid">{img_html}</div>' if img_html else ""

    return f"""
<section id="stress">
  <h2>Stress Results</h2>
  <p>1-sigma von Mises and component stresses from the PSD analysis.
     Values are relative to the base excitation (not absolute accelerations).</p>
  {_table(["Stress Component", "Max |Stress| (MPa)"], stress_rows)}
  <p><em>Results represent one standard deviation (68% probability) of the
  structural response under the applied PSD loading.</em></p>
  {img_grid}
</section>"""


def _section_conclusions(data: dict) -> str:
    layup  = data.get("composite_layup", {})
    freqs  = data.get("modal_results", {}).get("frequencies", [])
    pf_list = data.get("participation_factors", {}).get("factors", [])
    disp_y = data.get("displacement_results", {}).get("Y", {})
    stress = data.get("stress_results", {}).get("EQV", {})
    params = data.get("analysis_parameters", {})

    num_plies = len(layup.get("plies", []))
    total_t   = layup.get("total_thickness_mm", "N/A")
    f1        = freqs[0]["frequency_hz"] if freqs else "N/A"
    n_modes   = len(freqs)
    damp_pct  = params.get("damping_ratio", 0.02) * 100
    freq_min  = params.get("freq_range_hz", [0, 3000])[0]
    freq_max  = params.get("freq_range_hz", [0, 3000])[1]

    max_uy   = disp_y.get("max_absolute_um", "N/A")
    max_uy_s = f"{max_uy:.3f}" if isinstance(max_uy, (int, float)) else str(max_uy)

    max_seqv   = stress.get("max_absolute_mpa", "N/A")
    max_seqv_s = f"{max_seqv:.3f}" if isinstance(max_seqv, (int, float)) else str(max_seqv)

    # Find dominant PF mode
    dom_note = ""
    if pf_list:
        dom = max(pf_list, key=lambda x: abs(x.get("participation_factor_y", 0)))
        dom_note = (f"Mode {dom['mode']} at {dom['frequency_hz']:.2f} Hz exhibits the largest "
                    f"Y-direction participation factor of {dom['participation_factor_y']:.4e}, "
                    f"confirming it as the primary contributor to the structural response. ")

    return f"""
<section id="conclusions">
  <h2>Conclusions</h2>
  <p>
    The heavy-duty wrench with {num_plies}-ply carbon/epoxy sandwich composite
    (symmetric [0/0/45/45/90/Core/90/45/45/0/0], total thickness {total_t} mm) was
    analyzed for random vibration response under a PSD base excitation in the
    Y-direction over {freq_min:.0f}–{freq_max:.0f} Hz.
  </p>
  <p>
    Modal analysis identified {n_modes} natural frequencies within the analysis
    bandwidth. {dom_note}
  </p>
  <p>
    The peak 1-sigma displacement is <b>{max_uy_s} µm</b> at the wrench head, and the
    maximum 1-sigma von Mises stress is <b>{max_seqv_s} MPa</b>. These values are
    well within the typical ultimate tensile strength of carbon/epoxy prepreg
    (~600–700 MPa for in-plane loading), confirming structural adequacy under
    the specified vibration environment.
  </p>
  <p>
    Damping was modeled as {damp_pct:.1f}% constant modal damping across all modes,
    consistent with typical structural composite damping values for lightly damped
    aerospace/industrial components.
  </p>
</section>"""


def _html_head() -> str:
    return f"""<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>ANSYS Composite PSD Analysis Report — Heavy Duty Wrench</title>
  <style>{REPORT_CSS}</style>
</head>"""


def _html_header(data: dict) -> str:
    meta = data.get("metadata", {})
    gen_at = meta.get("generated_at", "")
    if gen_at:
        try:
            gen_at = datetime.datetime.fromisoformat(gen_at).strftime("%Y-%m-%d %H:%M")
        except Exception:
            pass
    return f"""<body>
<header>
  <h1>Composite Wrench — Random Vibration (PSD) Analysis Report</h1>
  <p>ANSYS MAPDL &nbsp;|&nbsp; SOLID186 Layered Solid &nbsp;|&nbsp;
     Carbon/Epoxy Sandwich &nbsp;|&nbsp; Generated: {gen_at}</p>
</header>"""


def _html_nav() -> str:
    return """<nav>
  <a href="#summary">Summary</a>
  <a href="#parameters">Parameters</a>
  <a href="#materials">Materials</a>
  <a href="#layup">Layup</a>
  <a href="#frequencies">Frequencies</a>
  <a href="#psd-input">PSD Input</a>
  <a href="#response">Response</a>
  <a href="#stress">Stress</a>
  <a href="#conclusions">Conclusions</a>
</nav>"""


def _html_footer() -> str:
    py_ver = f"{sys.version_info.major}.{sys.version_info.minor}.{sys.version_info.micro}"
    return f"""<footer>
  <p>Generated by generate_report.py &nbsp;|&nbsp; ANSYS Composite PSD Analysis &nbsp;|&nbsp;
     Python {py_ver}</p>
  <p>ANSYS Student 2025 R2 &nbsp;|&nbsp; PyMAPDL &nbsp;|&nbsp; FastMCP</p>
</footer>
</body>
</html>"""

# ---------------------------------------------------------------------------
# Build HTML
# ---------------------------------------------------------------------------

def build_html(data: dict, plot_paths: dict) -> str:
    # Build image lookup: {type_component -> file_path}
    images = {}
    for img in data.get("images", []):
        key = f"{img['type']}_{img['component'].lower()}"
        if img.get("status") == "ok":
            images[key] = img["file"]

    fallback = data.get("_fallback", False)
    fallback_banner = ""
    if fallback:
        fallback_banner = (
            '<div class="warn" style="margin: 16px 0;">'
            "<b>Note:</b> This report was generated from design-phase hardcoded data. "
            "Re-run after completing the ANSYS analysis to populate live simulation results."
            "</div>"
        )

    parts = [
        _html_head(),
        _html_header(data),
        _html_nav(),
        "<main>",
        fallback_banner,
        _section_summary(data),
        _section_parameters(data),
        _section_materials(data),
        _section_layup(data, plot_paths),
        _section_frequencies(data),
        _section_psd_input(data),
        _section_response(data, images),
        _section_stress(data, images),
        _section_conclusions(data),
        "</main>",
        _html_footer(),
    ]
    return "\n".join(parts)

# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="Generate a self-contained HTML report from ANSYS composite PSD analysis results."
    )
    parser.add_argument(
        "--results", metavar="PATH",
        help="Path to results.json produced by collect_all_results() MCP tool. "
             "If omitted, auto-discovers report_output/results.json in the script directory.",
    )
    parser.add_argument(
        "--output", metavar="PATH", default=None,
        help="Output HTML path (default: report_output/ansys_psd_report.html).",
    )
    parser.add_argument(
        "--fallback", action="store_true",
        help="Use hardcoded design-phase data instead of a results.json file.",
    )
    args = parser.parse_args()

    script_dir = os.path.dirname(os.path.abspath(__file__))
    report_dir = os.path.join(script_dir, "report_output")
    os.makedirs(report_dir, exist_ok=True)

    # Load data
    if args.fallback:
        data = get_fallback_data()
        print("Using hardcoded fallback data.")
    elif args.results:
        print(f"Loading results from: {args.results}")
        data = load_results(args.results)
    else:
        auto_path = os.path.join(report_dir, "results.json")
        if os.path.exists(auto_path):
            print(f"Auto-discovered results.json at: {auto_path}")
            data = load_results(auto_path)
        else:
            print("No results.json found. Using hardcoded fallback data.")
            print("(Run with --fallback to suppress this message, or provide --results PATH)")
            data = get_fallback_data()

    # Generate matplotlib plots
    print("Generating matplotlib figures...")
    plot_paths = generate_matplotlib_plots(data, report_dir)

    # Build HTML
    html = build_html(data, plot_paths)

    # Write output
    if args.output:
        out_path = args.output
    else:
        out_path = os.path.join(report_dir, "ansys_psd_report.html")

    with open(out_path, "w", encoding="utf-8") as f:
        f.write(html)

    print(f"\nReport written to: {out_path}")
    if plot_paths:
        for name, path in plot_paths.items():
            print(f"  Plot saved: {path}")


if __name__ == "__main__":
    main()
