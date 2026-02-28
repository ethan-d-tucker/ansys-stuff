"""
Material library for composite PSD analysis.

Provides elastic properties (for ANSYS) and strength allowables (for failure analysis)
for common composite materials, plus default layup definitions.
"""

# ---------------------------------------------------------------------------
# Material definitions
# ---------------------------------------------------------------------------

MATERIAL_LIBRARY = {
    "carbon_epoxy_woven": {
        "name": "Carbon/Epoxy Woven Prepreg (T300/Epoxy)",
        "description": (
            "Standard modulus carbon fiber woven fabric, 5-harness satin weave, "
            "epoxy matrix. Balanced weave gives EX ~ EY."
        ),
        # Elastic constants (SI: Pa, kg/m3) -- used by ANSYS
        "elastic": {
            "EX": 60.0e9,
            "EY": 60.0e9,
            "EZ": 10.0e9,
            "GXY": 5.0e9,
            "GXZ": 4.0e9,
            "GYZ": 4.0e9,
            "PRXY": 0.04,
            "PRXZ": 0.30,
            "PRYZ": 0.30,
            "DENS": 1420.0,
        },
        # Strength allowables (Pa) -- used by failure analysis
        "strength": {
            "Xt": 730e6,    # Tensile strength, fibre direction 1
            "Xc": 630e6,    # Compressive strength, fibre direction 1
            "Yt": 730e6,    # Tensile strength, transverse direction 2 (woven ~ symmetric)
            "Yc": 630e6,    # Compressive strength, transverse direction 2
            "S12": 70e6,    # In-plane shear strength
            "S13": 55e6,    # Interlaminar shear XZ
            "S23": 55e6,    # Interlaminar shear YZ
            "Zt": 50e6,     # Through-thickness tensile
            "Zc": 250e6,    # Through-thickness compressive
        },
    },
    "honeycomb_core": {
        "name": "Aramid Honeycomb Core (Nomex-type)",
        "description": (
            "Aramid fibre / phenolic resin honeycomb core, 48 kg/m3 density. "
            "High through-thickness stiffness and shear, negligible in-plane stiffness."
        ),
        "elastic": {
            "EX": 1.0e6,
            "EY": 1.0e6,
            "EZ": 130.0e6,
            "GXY": 1.0e6,
            "GXZ": 24.0e6,
            "GYZ": 48.0e6,
            "PRXY": 0.49,
            "PRXZ": 0.001,
            "PRYZ": 0.001,
            "DENS": 48.0,
        },
        "strength": {
            "Xt": 2.5e6,
            "Xc": 2.5e6,
            "Yt": 2.5e6,
            "Yc": 2.5e6,
            "S12": 1.5e6,
            "S13": 1.5e6,
            "S23": 1.5e6,
            "Zt": 3.5e6,
            "Zc": 10.0e6,
        },
    },
}

# ---------------------------------------------------------------------------
# Default composite layup  [0/0/45/45/90/90/45/45/0/0 / 90 / 0/0/45/45/90/90/45/45/0/0]
# 21-ply all-carbon symmetric laminate
# ---------------------------------------------------------------------------

DEFAULT_LAYUP = [
    {"ply": 1,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Outer"},
    {"ply": 2,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Laminate"},
    {"ply": 3,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 4,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 5,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 90, "role": "Laminate"},
    {"ply": 6,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 90, "role": "Laminate"},
    {"ply": 7,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 8,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 9,  "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Laminate"},
    {"ply": 10, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Laminate"},
    {"ply": 11, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 90, "role": "Center"},
    {"ply": 12, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Laminate"},
    {"ply": 13, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Laminate"},
    {"ply": 14, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 15, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 16, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 90, "role": "Laminate"},
    {"ply": 17, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 90, "role": "Laminate"},
    {"ply": 18, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 19, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 45, "role": "Laminate"},
    {"ply": 20, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Laminate"},
    {"ply": 21, "mat": "carbon_epoxy_woven", "thickness_mm": 0.15, "angle": 0,  "role": "Outer"},
]


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_material(name):
    """Return full material dict (elastic + strength) by library key."""
    if name not in MATERIAL_LIBRARY:
        raise KeyError(f"Unknown material '{name}'. Available: {list(MATERIAL_LIBRARY)}")
    return MATERIAL_LIBRARY[name]


def get_elastic_props(name):
    """Return the ANSYS-ready elastic property dict for *name*."""
    return get_material(name)["elastic"]


def get_strength_allowables(name):
    """Return strength allowable dict for *name*."""
    return get_material(name)["strength"]


def list_materials():
    """Return list of available material keys."""
    return list(MATERIAL_LIBRARY)


def get_default_layup():
    """Return a copy of the default layup list."""
    import copy
    return copy.deepcopy(DEFAULT_LAYUP)


def get_layup_summary(layup=None):
    """Return a summary dict for a layup (total thickness, sequence string, etc.)."""
    if layup is None:
        layup = DEFAULT_LAYUP
    total_mm = sum(p["thickness_mm"] for p in layup)
    angles = [str(int(p["angle"])) for p in layup]
    n_face = sum(1 for p in layup if "core" not in p["mat"].lower())
    n_core = sum(1 for p in layup if "core" in p["mat"].lower())
    return {
        "n_plies": len(layup),
        "n_face_plies": n_face,
        "n_core_plies": n_core,
        "total_thickness_mm": round(total_mm, 3),
        "stacking_sequence": "/".join(angles),
        "symmetric": _is_symmetric(layup),
    }


def _is_symmetric(layup):
    """Check if layup angles are symmetric about the midplane."""
    angles = [p["angle"] for p in layup]
    n = len(angles)
    for i in range(n // 2):
        if angles[i] != angles[n - 1 - i]:
            return False
    return True
