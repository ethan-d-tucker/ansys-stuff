"""
Composite failure analysis engine.

Computes Tsai-Wu and Max Stress failure indices from 1-sigma stress fields
produced by the PSD SRSS approach.  All computations are pure numpy -- no
ANSYS dependency.

Stress sign convention
----------------------
SRSS always produces positive magnitudes.  For the failure criteria we need
the sign (tension vs compression).  We adopt the *dominant-mode sign*
convention: the sign at each node comes from whichever mode contributes the
largest absolute stress at that node.
"""

from dataclasses import dataclass
import numpy as np

from material_library import get_strength_allowables


# ---------------------------------------------------------------------------
# Data classes
# ---------------------------------------------------------------------------

@dataclass
class PlyFailureResult:
    """Failure results for a single ply across all nodes."""
    ply_index: int
    angle_deg: float
    material: str
    tsai_wu_index: np.ndarray       # (n_nodes,)
    max_stress_index: np.ndarray    # (n_nodes,)


@dataclass
class FailureResult:
    """Aggregated failure results (envelope over all plies)."""
    # Per-node envelopes (worst ply at each node)
    tsai_wu_index: np.ndarray       # (n_nodes,)  max TW index across plies
    max_stress_index: np.ndarray    # (n_nodes,)  max MS index across plies
    tsai_wu_fos: np.ndarray         # (n_nodes,)  1 / TW index  (capped at 999)
    max_stress_fos: np.ndarray      # (n_nodes,)  1 / MS index  (capped at 999)
    critical_ply_tw: np.ndarray     # (n_nodes,)  ply number driving TW
    critical_ply_ms: np.ndarray     # (n_nodes,)  ply number driving MS

    # Scalar summaries
    max_tw_index: float
    max_ms_index: float
    min_fos_tw: float
    min_fos_ms: float
    overall_pass_tw: bool           # True if max TW index < 1.0
    overall_pass_ms: bool           # True if max MS index < 1.0

    # Per-ply detail
    ply_results: list               # list[PlyFailureResult]


# ---------------------------------------------------------------------------
# Stress rotation
# ---------------------------------------------------------------------------

def rotate_stress_to_ply(sx, sy, sxy, theta_deg):
    """
    Rotate in-plane stress (sx, sy, sxy) from global to ply coordinates.

    Parameters
    ----------
    sx, sy, sxy : array_like   Global stress components (Pa).
    theta_deg   : float        Ply angle in degrees.

    Returns
    -------
    s1, s2, t12 : arrays  Stress in ply (fibre / transverse / shear) coords.
    """
    theta = np.radians(theta_deg)
    c = np.cos(theta)
    s = np.sin(theta)
    c2 = c * c
    s2 = s * s
    cs = c * s

    s1  = c2 * sx + s2 * sy + 2.0 * cs * sxy
    s2_ = s2 * sx + c2 * sy - 2.0 * cs * sxy
    t12 = -cs * sx + cs * sy + (c2 - s2) * sxy
    return s1, s2_, t12


# ---------------------------------------------------------------------------
# Tsai-Wu single-point
# ---------------------------------------------------------------------------

def _tsai_wu_index(s1, s2, t12, Xt, Xc, Yt, Yc, S12):
    """
    Tsai-Wu failure index for a single stress state.

    F_TW = F1*s1 + F2*s2 + F11*s1^2 + F22*s2^2 + F66*t12^2 + 2*F12*s1*s2

    Returns failure index (>= 1.0 means failure).
    """
    F1  = 1.0 / Xt - 1.0 / Xc
    F2  = 1.0 / Yt - 1.0 / Yc
    F11 = 1.0 / (Xt * Xc)
    F22 = 1.0 / (Yt * Yc)
    F66 = 1.0 / (S12 * S12)
    F12 = -0.5 * np.sqrt(F11 * F22)  # standard interaction approximation

    index = (F1 * s1 + F2 * s2
             + F11 * s1**2 + F22 * s2**2
             + F66 * t12**2
             + 2.0 * F12 * s1 * s2)
    return index


# ---------------------------------------------------------------------------
# Max Stress single-point
# ---------------------------------------------------------------------------

def _max_stress_index(s1, s2, t12, Xt, Xc, Yt, Yc, S12):
    """
    Max Stress failure index.

    R = max( s1/Xt or |s1|/Xc,  s2/Yt or |s2|/Yc,  |t12|/S12 )
    """
    # Direction 1
    r1 = np.where(s1 >= 0, s1 / Xt, np.abs(s1) / Xc)
    # Direction 2
    r2 = np.where(s2 >= 0, s2 / Yt, np.abs(s2) / Yc)
    # Shear
    r12 = np.abs(t12) / S12

    return np.maximum(np.maximum(r1, r2), r12)


# ---------------------------------------------------------------------------
# Dominant-mode signed SRSS
# ---------------------------------------------------------------------------

def signed_srss(modal_components, modal_sigma2):
    """
    Compute 1-sigma magnitude via SRSS, with sign from the dominant mode.

    Parameters
    ----------
    modal_components : ndarray (n_modes, n_nodes)
        Modal stress/displacement component per mode per node.
    modal_sigma2 : ndarray (n_modes,)
        Variance of modal response for each mode.

    Returns
    -------
    signed_result : ndarray (n_nodes,)
        1-sigma values with sign from the dominant-contribution mode.
    """
    n_modes, n_nodes = modal_components.shape
    variance = np.zeros(n_nodes)
    for i in range(n_modes):
        variance += modal_components[i, :] ** 2 * modal_sigma2[i]
    magnitude = np.sqrt(variance)

    # Sign from mode with largest weighted contribution at each node
    weighted = np.abs(modal_components) * np.sqrt(modal_sigma2)[:, np.newaxis]
    dominant_mode = np.argmax(weighted, axis=0)
    signs = np.sign(modal_components[dominant_mode, np.arange(n_nodes)])
    signs[signs == 0] = 1.0

    return signs * magnitude


# ---------------------------------------------------------------------------
# Main entry point
# ---------------------------------------------------------------------------

def compute_failure_indices(
    stress_sx, stress_sy, stress_sxy,
    layup, materials_dict=None,
    required_fos=1.5,
):
    """
    Compute Tsai-Wu and Max Stress failure indices at every node for all plies.

    Parameters
    ----------
    stress_sx : ndarray (n_nodes,)
        1-sigma normal stress X (Pa), signed.
    stress_sy : ndarray (n_nodes,)
        1-sigma normal stress Y (Pa), signed.
    stress_sxy : ndarray (n_nodes,)
        1-sigma shear stress XY (Pa), signed.
    layup : list[dict]
        Ply definitions with keys: mat, angle, thickness_mm.
    materials_dict : dict or None
        Override material library.  If None, uses material_library module.
    required_fos : float
        Required factor of safety for pass/fail assessment.

    Returns
    -------
    FailureResult
    """
    n_nodes = len(stress_sx)
    sx = np.asarray(stress_sx, dtype=float)
    sy = np.asarray(stress_sy, dtype=float)
    sxy = np.asarray(stress_sxy, dtype=float)

    # Envelope arrays
    tw_envelope = np.full(n_nodes, -np.inf)
    ms_envelope = np.full(n_nodes, -np.inf)
    crit_ply_tw = np.zeros(n_nodes, dtype=int)
    crit_ply_ms = np.zeros(n_nodes, dtype=int)

    ply_results = []

    for ply in layup:
        mat_key = ply["mat"]
        angle = ply["angle"]
        ply_num = ply.get("ply", len(ply_results) + 1)

        # Get strength allowables
        if materials_dict and mat_key in materials_dict:
            strength = materials_dict[mat_key].get(
                "strength", get_strength_allowables(mat_key)
            )
        else:
            strength = get_strength_allowables(mat_key)

        Xt  = strength["Xt"]
        Xc  = strength["Xc"]
        Yt  = strength["Yt"]
        Yc  = strength["Yc"]
        S12_val = strength["S12"]

        # Rotate to ply coordinates
        s1, s2, t12 = rotate_stress_to_ply(sx, sy, sxy, angle)

        # Failure indices
        tw_idx = _tsai_wu_index(s1, s2, t12, Xt, Xc, Yt, Yc, S12_val)
        ms_idx = _max_stress_index(s1, s2, t12, Xt, Xc, Yt, Yc, S12_val)

        ply_results.append(PlyFailureResult(
            ply_index=ply_num,
            angle_deg=angle,
            material=mat_key,
            tsai_wu_index=tw_idx,
            max_stress_index=ms_idx,
        ))

        # Update envelopes
        update_tw = tw_idx > tw_envelope
        tw_envelope[update_tw] = tw_idx[update_tw]
        crit_ply_tw[update_tw] = ply_num

        update_ms = ms_idx > ms_envelope
        ms_envelope[update_ms] = ms_idx[update_ms]
        crit_ply_ms[update_ms] = ply_num

    # Factors of safety  (cap at 999 to avoid inf for zero-stress nodes)
    with np.errstate(divide="ignore", invalid="ignore"):
        tw_fos = np.where(tw_envelope > 1e-12, 1.0 / tw_envelope, 999.0)
        tw_fos = np.minimum(tw_fos, 999.0)
        ms_fos = np.where(ms_envelope > 1e-12, 1.0 / ms_envelope, 999.0)
        ms_fos = np.minimum(ms_fos, 999.0)

    max_tw = float(np.max(tw_envelope))
    max_ms = float(np.max(ms_envelope))

    return FailureResult(
        tsai_wu_index=tw_envelope,
        max_stress_index=ms_envelope,
        tsai_wu_fos=tw_fos,
        max_stress_fos=ms_fos,
        critical_ply_tw=crit_ply_tw,
        critical_ply_ms=crit_ply_ms,
        max_tw_index=max_tw,
        max_ms_index=max_ms,
        min_fos_tw=float(np.min(tw_fos)),
        min_fos_ms=float(np.min(ms_fos)),
        overall_pass_tw=(max_tw < 1.0),
        overall_pass_ms=(max_ms < 1.0),
        ply_results=ply_results,
    )


def compute_core_failure(stress_sz, stress_sxz, stress_syz, core_mat_key="honeycomb_core"):
    """
    Check honeycomb core against through-thickness and transverse-shear allowables.

    Returns dict with indices, FoS, and pass/fail.
    """
    strength = get_strength_allowables(core_mat_key)
    sz  = np.asarray(stress_sz, dtype=float)
    sxz = np.asarray(stress_sxz, dtype=float)
    syz = np.asarray(stress_syz, dtype=float)

    # Through-thickness
    rz = np.where(sz >= 0, sz / strength["Zt"], np.abs(sz) / strength["Zc"])
    # Transverse shear
    rxz = np.abs(sxz) / strength["S13"]
    ryz = np.abs(syz) / strength["S23"]

    envelope = np.maximum(np.maximum(rz, rxz), ryz)
    with np.errstate(divide="ignore", invalid="ignore"):
        fos = np.where(envelope > 1e-12, 1.0 / envelope, 999.0)
        fos = np.minimum(fos, 999.0)

    return {
        "max_index": float(np.max(envelope)),
        "min_fos": float(np.min(fos)),
        "pass": bool(np.max(envelope) < 1.0),
        "index_array": envelope,
        "fos_array": fos,
    }
