"""Generate result plots for the composite wrench PSD analysis."""
import os
import matplotlib.pyplot as plt
import matplotlib.patches as mpatches
import numpy as np

# --- Natural Frequencies ---
modes = [1, 2, 3, 4, 5, 6]
freqs = [112.7096, 208.3092, 742.3481, 1374.2027, 1656.5302, 2128.7873]

# --- PSD Input Table ---
psd_freq = [20, 80, 350, 2000]
psd_val = [0.01, 0.04, 0.04, 0.007]

# --- Response PSD at Node 1500 (UY) from POST26 RPSD ---
# Extracted from PRVAR output - key points around resonances
rpsd_freq = [
    20.0, 50.56, 80.0, 100.15, 105.71, 108.9, 110.96, 112.29,
    112.59, 112.71, 113.13, 114.48, 118.13, 124.76, 150.55,
    200.10, 207.24, 208.31, 209.38, 220.07, 300.0, 350.0,
    500.0, 650.0, 700.0, 720.86, 733.95, 740.89, 742.35,
    744.71, 756.06, 800.0, 1000.0, 1200.0, 1374.2, 1500.0,
    1600.6, 1652.7, 1656.5, 1660.4, 1700.0, 1900.0, 2000.0
]
rpsd_val = [
    4.19e-9, 4.18e-10, 3.08e-10, 8.16e-10, 1.97e-9, 4.82e-9,
    1.08e-8, 1.58e-8, 1.60e-8, 1.60e-8, 1.52e-8, 9.19e-9,
    1.93e-9, 3.59e-10, 1.84e-11, 1.31e-12, 1.00e-12, 9.66e-13,
    9.31e-13, 6.52e-13, 7.48e-14, 3.60e-14, 1.26e-14, 1.55e-14,
    4.98e-14, 1.13e-13, 2.39e-13, 2.94e-13, 2.93e-13, 2.76e-13,
    1.41e-13, 1.18e-14, 1.24e-16, 1.52e-17, 4.91e-19, 1.01e-17,
    1.86e-16, 1.04e-15, 1.08e-15, 1.09e-15, 7.12e-16, 1.08e-16,
    1.63e-16
]

# --- Composite Layup ---
ply_angles = [0, 0, 45, 45, 90, 'core', 90, 45, 45, 0, 0]
ply_thicknesses = [0.15, 0.15, 0.15, 0.15, 0.15, 1.675, 0.15, 0.15, 0.15, 0.15, 0.15]  # mm
ply_materials = ['Carbon', 'Carbon', 'Carbon', 'Carbon', 'Carbon',
                 'Honeycomb', 'Carbon', 'Carbon', 'Carbon', 'Carbon', 'Carbon']

# --- Participation Factors ---
pf_modes = [1, 2, 3, 4, 5, 6]
pf_values = [-0.039861, 7.61e-11, 0.021524, 2.26e-13, 0.003453, -0.01220]

# ============== CREATE FIGURE ==============
fig = plt.figure(figsize=(16, 12))
fig.suptitle('Composite Wrench — Random Vibration (PSD) Analysis Results',
             fontsize=16, fontweight='bold', y=0.98)

# --- Plot 1: Natural Frequencies ---
ax1 = fig.add_subplot(2, 2, 1)
bars = ax1.bar(modes, freqs, color=['#2196F3', '#4CAF50', '#FF9800', '#E91E63', '#9C27B0', '#00BCD4'],
               edgecolor='black', linewidth=0.5)
for bar, f in zip(bars, freqs):
    ax1.text(bar.get_x() + bar.get_width()/2, bar.get_height() + 30,
             f'{f:.1f}', ha='center', va='bottom', fontsize=9, fontweight='bold')
ax1.set_xlabel('Mode Number')
ax1.set_ylabel('Natural Frequency (Hz)')
ax1.set_title('Natural Frequencies (6 Modes, 0-3000 Hz)')
ax1.set_ylim(0, max(freqs)*1.15)
ax1.grid(axis='y', alpha=0.3)

# --- Plot 2: PSD Input + Response PSD ---
ax2 = fig.add_subplot(2, 2, 2)
ax2.loglog(psd_freq, psd_val, 'b-o', linewidth=2, markersize=6, label='Input PSD (G²/Hz)', zorder=5)

# Add vertical lines at natural frequencies
for i, f in enumerate(freqs):
    if f <= 2000:
        ax2.axvline(x=f, color='red', alpha=0.3, linestyle='--', linewidth=0.8)
        if i < 3:  # Label first 3
            ax2.text(f, max(psd_val)*1.5, f'f{i+1}={f:.0f}Hz',
                     rotation=90, va='bottom', ha='right', fontsize=7, color='red')

ax2.set_xlabel('Frequency (Hz)')
ax2.set_ylabel('PSD (G²/Hz)')
ax2.set_title('PSD Input Spectrum with Natural Frequencies')
ax2.legend(loc='upper right')
ax2.grid(True, which='both', alpha=0.3)
ax2.set_xlim(10, 3000)

# --- Plot 3: Response PSD at Node 1500 ---
ax3 = fig.add_subplot(2, 2, 3)
rpsd_freq_np = np.array(rpsd_freq)
rpsd_val_np = np.array(rpsd_val)
# Filter out zeros for log plot
mask = rpsd_val_np > 0
ax3.loglog(rpsd_freq_np[mask], rpsd_val_np[mask], 'r-', linewidth=1.5, label='Response PSD (Node 1500, UY)')
for i, f in enumerate(freqs):
    if f <= 2000:
        ax3.axvline(x=f, color='blue', alpha=0.2, linestyle='--', linewidth=0.8)
ax3.set_xlabel('Frequency (Hz)')
ax3.set_ylabel('Response PSD (m²/Hz)')
ax3.set_title('Response PSD — Node 1500, UY Direction')
ax3.legend(loc='upper right')
ax3.grid(True, which='both', alpha=0.3)
ax3.set_xlim(10, 3000)

# Annotate peak
peak_idx = np.argmax(rpsd_val_np)
ax3.annotate(f'Peak at {rpsd_freq_np[peak_idx]:.1f} Hz\n(Mode 1 resonance)',
             xy=(rpsd_freq_np[peak_idx], rpsd_val_np[peak_idx]),
             xytext=(300, rpsd_val_np[peak_idx]*0.5),
             arrowprops=dict(arrowstyle='->', color='black'),
             fontsize=8, ha='center')

# --- Plot 4: Composite Layup Visualization ---
ax4 = fig.add_subplot(2, 2, 4)
colors_map = {'Carbon': '#333333', 'Honeycomb': '#FFD700'}
y_pos = 0
for i, (angle, t, mat) in enumerate(zip(ply_angles, ply_thicknesses, ply_materials)):
    color = colors_map[mat]
    rect = mpatches.FancyBboxPatch((0.2, y_pos), 0.6, t,
                                     boxstyle="round,pad=0.01",
                                     facecolor=color, edgecolor='white', linewidth=1)
    ax4.add_patch(rect)
    label = f'{angle}°' if isinstance(angle, int) else 'Core'
    text_color = 'white' if mat == 'Carbon' else 'black'
    ax4.text(0.5, y_pos + t/2, f'{label}  ({t:.3f} mm, {mat})',
             ha='center', va='center', fontsize=8, fontweight='bold', color=text_color)
    y_pos += t

ax4.set_xlim(0, 1)
ax4.set_ylim(-0.1, y_pos + 0.1)
ax4.set_ylabel('Thickness (mm)')
ax4.set_title(f'Composite Sandwich Layup — Total: {sum(ply_thicknesses):.3f} mm')
ax4.set_xticks([])
ax4.set_aspect('auto')

# Add legend
carbon_patch = mpatches.Patch(facecolor='#333333', edgecolor='black', label='Carbon/Epoxy Prepreg')
honey_patch = mpatches.Patch(facecolor='#FFD700', edgecolor='black', label='Honeycomb Core')
ax4.legend(handles=[carbon_patch, honey_patch], loc='upper left', fontsize=8)

plt.tight_layout(rect=[0, 0, 1, 0.95])
output_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'psd_analysis_results.png')
plt.savefig(output_path, dpi=150, bbox_inches='tight')
plt.close()
print(f"Plot saved to: {output_path}")

# --- Compute approximate 1-sigma values by integrating response PSD ---
# Trapezoidal integration of RPSD over frequency
variance = np.trapezoid(rpsd_val_np[mask], rpsd_freq_np[mask])
rms_uy = np.sqrt(variance)
print(f"\nApproximate 1-sigma UY displacement at Node 1500:")
print(f"  Variance: {variance:.6e} m²")
print(f"  1-sigma (RMS): {rms_uy:.6e} m = {rms_uy*1000:.6f} mm = {rms_uy*1e6:.3f} um")

# --- Summary Table ---
print(f"\n{'='*60}")
print(f"  COMPOSITE WRENCH PSD ANALYSIS SUMMARY")
print(f"{'='*60}")
print(f"  Model: Heavy Duty Wrench (Parasolid)")
print(f"  Elements: 228 (SOLID186 layered hex/wedge)")
print(f"  Nodes: 1643")
print(f"  Composite: [0/0/45/45/90/core/90/45/45/0/0]")
print(f"  Total thickness: 3.175 mm")
print(f"  Materials: Carbon/Epoxy prepreg + Honeycomb core")
print(f"  BCs: Fixed handle end (X=0 to 5mm), 188 nodes")
print(f"  PSD Input: 20-2000 Hz, peak 0.04 G²/Hz")
print(f"  Damping: 2% constant modal")
print(f"{'='*60}")
print(f"  NATURAL FREQUENCIES:")
for m, f in zip(modes, freqs):
    pf = pf_values[m-1]
    print(f"    Mode {m}: {f:10.2f} Hz  (PF = {pf:+.4e})")
print(f"{'='*60}")
print(f"  1-SIGMA RESPONSE (Node 1500, wrench head):")
print(f"    UY displacement: {rms_uy*1e6:.3f} um")
print(f"    Dominant mode: Mode 1 at 112.7 Hz")
print(f"{'='*60}")
