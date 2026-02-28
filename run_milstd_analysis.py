"""
MIL-STD-810H Multi-Environment Composite PSD Analysis
======================================================
Top-level orchestrator that:
  1. Imports geometry (any Parasolid / STEP file)
  2. Runs modal analysis once
  3. Computes PSD response for each MIL-STD-810H environment x axis
  4. Evaluates composite failure (Tsai-Wu + Max Stress)
  5. Generates a professional DOCX report

Batch mode:
    python run_milstd_analysis.py --geometry WrenchParasolid.x_t \\
        --part-name "Heavy-Duty Wrench" \\
        --profiles MIN_INTEGRITY,HELICOPTER,JET_AIRCRAFT

Interactive mode: driven by Claude via MCP tools.
"""

import argparse
import os
import sys
import time

import numpy as np

from material_library import (
    get_material, get_elastic_props, get_default_layup, get_layup_summary,
)
from mil_std_profiles import get_all_profiles, get_profile, get_profile_names
from composite_failure import compute_failure_indices, compute_core_failure
from simulation_engine import SimulationConfig, run_multi_environment
from milstd_report import generate_milstd_report


SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))


def build_config(args):
    """Build SimulationConfig from CLI args."""
    geo_path = args.geometry
    if not os.path.isabs(geo_path):
        geo_path = os.path.join(SCRIPT_DIR, geo_path)

    material_name = args.material
    elastic = get_elastic_props(material_name)

    return SimulationConfig(
        geometry_file=geo_path,
        part_name=args.part_name,
        element_size=args.element_size,
        material_props=elastic,
        material_name=get_material(material_name)["name"],
        psd_table=[],   # overridden per profile
        excitation_direction="Y",  # overridden per axis
        damping_ratio=args.damping,
        num_modes=args.num_modes,
        freq_start=0.0,
        freq_end=args.freq_end,
    )


def run_analysis(config, profile_ids, layup, material_key, required_fos=1.5,
                 output_dir=None):
    """
    Run the full MIL-STD-810H analysis pipeline.

    Parameters
    ----------
    config : SimulationConfig
    profile_ids : list[str]  e.g. ["MIN_INTEGRITY", "HELICOPTER"]
    layup : list[dict]  from material_library
    material_key : str  key into MATERIAL_LIBRARY
    required_fos : float
    output_dir : str or None

    Returns
    -------
    dict with report_path, overall results, etc.
    """
    if output_dir is None:
        output_dir = os.path.join(SCRIPT_DIR, "report_output")
    os.makedirs(output_dir, exist_ok=True)

    t0 = time.time()

    # Load profiles
    profiles = [get_profile(pid) for pid in profile_ids]
    print(f"\n{'='*60}")
    print(f"  MIL-STD-810H PSD Analysis: {config.part_name}")
    print(f"  Environments: {', '.join(profile_ids)}")
    print(f"  Geometry: {os.path.basename(config.geometry_file)}")
    print(f"{'='*60}\n")

    # ---- Step 1-8: Simulation (modal once, PSD per env x axis) ----
    print("Phase 1: Simulation ...")
    mapdl, modal_data, env_results = run_multi_environment(
        config, profiles, axes=["X", "Y", "Z"]
    )

    # ---- Step 9: Composite failure analysis ----
    print("\nPhase 2: Composite failure analysis ...")
    material_info = get_material(material_key)
    failure_data = {}       # {(pid, axis): FailureResult}
    core_failure = {}       # {(pid, axis): dict}

    # Use free-node mask to exclude stress singularities at fixed BCs
    free = modal_data.free_mask

    # Check if layup contains core material (for core failure check)
    has_core = any("core" in p["mat"].lower() for p in layup)

    for pid, axes_data in env_results.items():
        for axis, psd_res in axes_data.items():
            label = f"  {pid}/{axis}: "

            # Tsai-Wu / Max Stress (free nodes only -- avoids BC singularities)
            fr = compute_failure_indices(
                stress_sx=psd_res.stress_sx[free],
                stress_sy=psd_res.stress_sy[free],
                stress_sxy=psd_res.stress_sxy[free],
                layup=layup,
                required_fos=required_fos,
            )
            failure_data[(pid, axis)] = fr
            print(f"{label}TW={fr.max_tw_index:.6f} (FoS={fr.min_fos_tw:.1f}), "
                  f"MS={fr.max_ms_index:.6f} (FoS={fr.min_fos_ms:.1f})")

            # Core failure (only if layup includes core material)
            if has_core:
                cf = compute_core_failure(
                    psd_res.stress_sz[free], psd_res.stress_sxz[free],
                    psd_res.stress_syz[free],
                    core_mat_key="honeycomb_core",
                )
            else:
                cf = {"max_index": 0.0, "min_fos": 999.0, "pass": True}
            core_failure[(pid, axis)] = cf

    # ---- Step 10: Generate report ----
    print("\nPhase 3: Generating report ...")
    report_path = generate_milstd_report(
        modal_data=modal_data,
        env_results=env_results,
        failure_data=failure_data,
        core_failure_data=core_failure,
        profiles=profiles,
        layup=layup,
        material_info=material_info,
        config=config,
        output_dir=output_dir,
        required_fos=required_fos,
    )

    # ---- Cleanup ----
    try:
        mapdl.finish()
        mapdl.exit()
    except Exception:
        pass

    elapsed = time.time() - t0

    # ---- Summary ----
    min_fos_tw = min(fr.min_fos_tw for fr in failure_data.values())
    min_fos_ms = min(fr.min_fos_ms for fr in failure_data.values())
    overall_pass = min_fos_tw >= required_fos and min_fos_ms >= required_fos

    print(f"\n{'='*60}")
    print(f"  ANALYSIS COMPLETE")
    print(f"{'='*60}")
    print(f"  Environments tested : {len(profiles)}")
    print(f"  Axes per env        : 3 (X, Y, Z)")
    print(f"  Total PSD cases     : {len(profiles) * 3}")
    print(f"  Min FoS (Tsai-Wu)   : {min_fos_tw:.1f}")
    print(f"  Min FoS (Max Stress): {min_fos_ms:.1f}")
    print(f"  Required FoS        : {required_fos:.1f}")
    print(f"  Overall result      : {'PASS' if overall_pass else 'FAIL'}")
    print(f"  Report              : {report_path}")
    print(f"  Total time          : {elapsed:.1f}s")

    return {
        "report_path": report_path,
        "overall_pass": overall_pass,
        "min_fos_tw": min_fos_tw,
        "min_fos_ms": min_fos_ms,
        "n_cases": len(profiles) * 3,
        "elapsed_s": elapsed,
    }


# ---------------------------------------------------------------------------
# CLI entry point
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(
        description="MIL-STD-810H Multi-Environment Composite PSD Analysis"
    )
    parser.add_argument(
        "--geometry", required=True,
        help="Path to geometry file (Parasolid .x_t or STEP .stp)"
    )
    parser.add_argument(
        "--part-name", default="Composite Part",
        help="Human-readable part name for the report"
    )
    parser.add_argument(
        "--profiles", default="MIN_INTEGRITY,HELICOPTER,JET_AIRCRAFT",
        help="Comma-separated MIL-STD profile IDs "
             f"(available: {', '.join(get_profile_names())})"
    )
    parser.add_argument(
        "--material", default="carbon_epoxy_woven",
        help="Material library key"
    )
    parser.add_argument(
        "--element-size", type=float, default=0.003,
        help="Mesh element size in metres (default: 0.003 = 3mm)"
    )
    parser.add_argument(
        "--damping", type=float, default=0.02,
        help="Modal damping ratio (default: 0.02 = 2%%)"
    )
    parser.add_argument(
        "--num-modes", type=int, default=20,
        help="Number of modes to extract (default: 20)"
    )
    parser.add_argument(
        "--freq-end", type=float, default=3000.0,
        help="Upper frequency bound in Hz (default: 3000)"
    )
    parser.add_argument(
        "--required-fos", type=float, default=1.5,
        help="Required factor of safety (default: 1.5)"
    )
    parser.add_argument(
        "--output", default=None,
        help="Output directory (default: report_output/)"
    )

    args = parser.parse_args()

    config = build_config(args)
    profile_ids = [p.strip() for p in args.profiles.split(",")]
    layup = get_default_layup()

    result = run_analysis(
        config=config,
        profile_ids=profile_ids,
        layup=layup,
        material_key=args.material,
        required_fos=args.required_fos,
        output_dir=args.output,
    )

    sys.exit(0 if result["overall_pass"] else 1)


if __name__ == "__main__":
    main()
