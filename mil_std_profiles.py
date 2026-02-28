"""
MIL-STD-810H Method 514.8 vibration environment profiles.

Each profile defines a random vibration PSD table (g^2/Hz vs Hz),
requirement metadata, and test parameters.  Profile breakpoints are
representative of the standard's category envelopes.
"""

import numpy as np


# ---------------------------------------------------------------------------
# Profile definitions
# ---------------------------------------------------------------------------

PROFILES = {
    "MIN_INTEGRITY": {
        "id": "MIN_INTEGRITY",
        "name": "General Minimum Integrity",
        "mil_std_ref": "MIL-STD-810H Method 514.8, Category 24, Annex E",
        "requirement_id": "REQ-VIB-001",
        "description": (
            "Basic qualification vibration environment for general materiel "
            "not assigned to a specific platform.  Flat broadband profile "
            "representing minimum survivability requirement."
        ),
        "duration_min_per_axis": 60,
        "test_axes": ["X", "Y", "Z"],
        "psd_table": [
            (10.0,  0.015),
            (20.0,  0.015),
            (80.0,  0.040),
            (350.0, 0.040),
            (500.0, 0.015),
        ],
    },
    "TRUCK_TRANSPORT": {
        "id": "TRUCK_TRANSPORT",
        "name": "US Highway Truck Transport",
        "mil_std_ref": "MIL-STD-810H Method 514.8, Category 4, Annex A, Figure 514.8A-1",
        "requirement_id": "REQ-VIB-002",
        "description": (
            "Composite wheeled-vehicle vibration for US highway truck transport.  "
            "Characterised by higher low-frequency content from road surface "
            "roughness and suspension dynamics."
        ),
        "duration_min_per_axis": 60,
        "test_axes": ["X", "Y", "Z"],
        "psd_table": [
            (5.0,   0.001),
            (10.0,  0.010),
            (40.0,  0.010),
            (500.0, 0.001),
        ],
    },
    "HELICOPTER": {
        "id": "HELICOPTER",
        "name": "Helicopter Vibration",
        "mil_std_ref": "MIL-STD-810H Method 514.8, Category 9, Annex D",
        "requirement_id": "REQ-VIB-003",
        "description": (
            "Helicopter vibration environment combining broadband random from "
            "turbine engines, main/tail rotor harmonics, and aerodynamic buffet.  "
            "Higher overall levels than ground transport."
        ),
        "duration_min_per_axis": 120,
        "test_axes": ["X", "Y", "Z"],
        "psd_table": [
            (10.0,  0.005),
            (40.0,  0.020),
            (500.0, 0.020),
            (2000.0, 0.005),
        ],
    },
    "JET_AIRCRAFT": {
        "id": "JET_AIRCRAFT",
        "name": "Jet Aircraft Vibration",
        "mil_std_ref": "MIL-STD-810H Method 514.8, Category 7, Annex C, Figure 514.8C-1",
        "requirement_id": "REQ-VIB-004",
        "description": (
            "Fixed-wing jet aircraft vibration environment.  Broadband random "
            "from jet engine acoustic excitation and boundary-layer turbulence.  "
            "Highest severity in this profile set."
        ),
        "duration_min_per_axis": 120,
        "test_axes": ["X", "Y", "Z"],
        "psd_table": [
            (15.0,  0.010),
            (80.0,  0.040),
            (350.0, 0.040),
            (2000.0, 0.007),
        ],
    },
}


# ---------------------------------------------------------------------------
# Public API
# ---------------------------------------------------------------------------

def get_all_profiles():
    """Return dict of all MIL-STD-810H vibration profiles."""
    import copy
    return copy.deepcopy(PROFILES)


def get_profile(profile_id):
    """Return a single profile by ID."""
    if profile_id not in PROFILES:
        raise KeyError(
            f"Unknown profile '{profile_id}'. "
            f"Available: {list(PROFILES)}"
        )
    import copy
    return copy.deepcopy(PROFILES[profile_id])


def get_profile_names():
    """Return ordered list of profile IDs."""
    return list(PROFILES)


def compute_grms(psd_table):
    """
    Compute overall Grms from a PSD table using trapezoidal integration
    on a log-log interpolated grid.

    Parameters
    ----------
    psd_table : list of (freq_hz, g2_per_hz) tuples

    Returns
    -------
    float  Grms value
    """
    freqs = np.array([f for f, _ in psd_table])
    vals = np.array([v for _, v in psd_table])

    f_grid = np.logspace(np.log10(freqs[0]), np.log10(freqs[-1]), 2000)
    log_vals = np.interp(np.log10(f_grid), np.log10(freqs), np.log10(vals))
    psd_interp = 10.0 ** log_vals

    area = np.trapezoid(psd_interp, f_grid)
    return float(np.sqrt(area))


def psd_table_to_dicts(psd_table):
    """Convert [(freq, val), ...] to [{"frequency_hz": f, "psd_g2_per_hz": v}, ...]."""
    return [{"frequency_hz": f, "psd_g2_per_hz": v} for f, v in psd_table]


# ---------------------------------------------------------------------------
# Attach computed Grms to each profile on import
# ---------------------------------------------------------------------------
for _pid, _prof in PROFILES.items():
    _prof["grms"] = round(compute_grms(_prof["psd_table"]), 2)
