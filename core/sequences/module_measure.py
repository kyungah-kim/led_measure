"""Module measurement sequence.

Measures:
  1. Gamma: W/R/G/B channel gamma through configurable gray-level steps
  2. White: CCT, Duv (spec check data only; caller does PASS/FAIL)
  3. Color patches: R, G, B, C, M, Y, W  (full-field 100%)
  4. Gamut overlap stats  (DCI-P3, BT.2020) derived from R/G/B
  5. Color-accuracy delta: Δu'v' × 10 000 per patch vs. reference u'v'
"""

from __future__ import annotations

import math
from typing import TYPE_CHECKING, Callable, Dict, List, Optional, Tuple

from ..equipment.base import MeasureResult, PatternConfig, PatternInfo

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# ---------------------------------------------------------------------------
# Default gray-level steps for gamma measurement
# ---------------------------------------------------------------------------
DEFAULT_GAMMA_STEPS: List[int] = [0, 16, 32, 64, 96, 128, 160, 192, 224, 255]

# ---------------------------------------------------------------------------
# Color patches: (name, R, G, B) — displayed in this order
# ---------------------------------------------------------------------------
_COLOR_PATCHES: List[Tuple[str, int, int, int]] = [
    ("R",   255,   0,   0),
    ("G",     0, 255,   0),
    ("B",     0,   0, 255),
    ("C",     0, 255, 255),
    ("M",   255,   0, 255),
    ("Y",   255, 255,   0),
    ("W",   255, 255, 255),
]

# ---------------------------------------------------------------------------
# BT.709 / sRGB reference u'v' (CIE 1976 UCS) — used as default targets
# ---------------------------------------------------------------------------
# Derived from BT.709 primaries (xy) and D65 white point
def _xy_to_uv(x: float, y: float) -> Tuple[float, float]:
    d = -2.0 * x + 12.0 * y + 3.0
    if abs(d) < 1e-12:
        return 0.0, 0.0
    return 4.0 * x / d, 9.0 * y / d


# BT.709 primaries → linear XYZ (D65 normalised to Y=1 for full white)
_BT709_RGB_TO_XYZ = [
    [0.4124564, 0.3575761, 0.1804375],
    [0.2126729, 0.7151522, 0.0721750],
    [0.0193339, 0.1191920, 0.9503041],
]


def _rgb_to_xyz(r: float, g: float, b: float) -> Tuple[float, float, float]:
    m = _BT709_RGB_TO_XYZ
    X = m[0][0] * r + m[0][1] * g + m[0][2] * b
    Y = m[1][0] * r + m[1][1] * g + m[1][2] * b
    Z = m[2][0] * r + m[2][1] * g + m[2][2] * b
    return X, Y, Z


def _xyz_to_xy(X: float, Y: float, Z: float) -> Tuple[float, float]:
    s = X + Y + Z
    if s < 1e-12:
        return 0.0, 0.0
    return X / s, Y / s


def _patch_uv(r_lin: float, g_lin: float, b_lin: float) -> Tuple[float, float]:
    X, Y, Z = _rgb_to_xyz(r_lin, g_lin, b_lin)
    x, y = _xyz_to_xy(X, Y, Z)
    return _xy_to_uv(x, y)


BT709_REF_UV: Dict[str, Tuple[float, float]] = {
    "R": _patch_uv(1.0, 0.0, 0.0),
    "G": _patch_uv(0.0, 1.0, 0.0),
    "B": _patch_uv(0.0, 0.0, 1.0),
    "C": _patch_uv(0.0, 1.0, 1.0),
    "M": _patch_uv(1.0, 0.0, 1.0),
    "Y": _patch_uv(1.0, 1.0, 0.0),
    "W": _xy_to_uv(0.3127, 0.3290),   # D65
}

# ---------------------------------------------------------------------------
# Gamma calculation helper
# ---------------------------------------------------------------------------

def calc_gamma(levels: List[int], lv_values: List[float]) -> List[Optional[float]]:
    """Return per-point gamma values (None for level 0 or when undefined).

    gamma = log(L / L_max) / log(level / 255)

    L_max is taken as the luminance at level=255 (last point if present).
    """
    if not levels or not lv_values or len(levels) != len(lv_values):
        return []

    lv_max: Optional[float] = None
    for lvl, lv in zip(levels, lv_values):
        if lvl == 255:
            lv_max = lv
            break

    gammas: List[Optional[float]] = []
    for lvl, lv in zip(levels, lv_values):
        if lvl == 0 or lvl == 255:
            gammas.append(None)
            continue
        if lv_max is None or lv_max <= 0 or lv <= 0:
            gammas.append(None)
            continue
        try:
            g = math.log(lv / lv_max) / math.log(lvl / 255.0)
            gammas.append(round(g, 4))
        except (ValueError, ZeroDivisionError):
            gammas.append(None)
    return gammas


# ---------------------------------------------------------------------------
# Sequence class
# ---------------------------------------------------------------------------

class ModuleMeasureSequence:
    """Module-level optical measurement sequence.

    Step order:
      1.  Gamma: for each selected channel (W/R/G/B), sweep gray levels.
      2.  Color patches: R, G, B, C, M, Y, W (full-field 100%).
      3.  Compute gamut overlap and Δu'v' vs. reference.

    Progress callback signature: callback(step_name: str, data: dict)
      step_name ∈ {"gamma", "color", "done"}
    """

    def __init__(self, engine: "MeasurementEngine") -> None:
        self._engine = engine
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def run(
        self,
        is_hdr: bool,
        gamma_channels: List[str],
        gamma_steps: List[int],
        ref_uv: Dict[str, Tuple[float, float]],
        callback: Callable[[str, dict], None],
        run_gamma: bool = True,
        run_colors: bool = True,
    ) -> Dict[str, object]:
        """Run the full module measurement sequence.

        Parameters
        ----------
        is_hdr:          HDR output mode.
        gamma_channels:  Subset of ["W", "R", "G", "B"] to measure.
        gamma_steps:     Gray levels 0–255 to test per channel.
        ref_uv:          Reference u'v' per color name for delta calculation.
        callback:        Called on every measured data point.

        Returns
        -------
        {
          "gamma":       {ch: [{"level": int, "Lv": float, "gamma": float|None,
                                "result": MeasureResult}, ...]},
          "colors":      {"R": MeasureResult, "G": ..., ...},
          "gamut_stats": {...}    # from calc_gamut_stats
          "delta_uv":    {"R": float, ...}   # Δu'v' × 10 000
        }
        """
        from ..gamut_utils import calc_gamut_stats

        gen = self._engine.generator
        meter = self._engine.meter
        if gen is None or not gen.is_connected:
            raise RuntimeError("Generator is not connected")
        if meter is None or not meter.is_connected:
            raise RuntimeError("Meter is not connected")

        if is_hdr:
            gen.set_hdr(True)
        else:
            gen.set_sdr()

        self._stop_requested = False

        # ── 1. Gamma measurement ────────────────────────────────────────
        gamma_data: Dict[str, List[dict]] = {}

        for ch in (gamma_channels if run_gamma else []):
            if self._stop_requested:
                break
            ch_points: List[dict] = []
            for level in gamma_steps:
                if self._stop_requested:
                    break
                r = level if ch in ("W", "R") else 0
                g = level if ch in ("W", "G") else 0
                b = level if ch in ("W", "B") else 0
                cfg = PatternConfig(
                    type="full_field", color=f"gamma_{ch}_{level}",
                    r=r, g=g, b=b,
                )
                gen.set_pattern(cfg)
                with self._engine.meter_lock:
                    mres = meter.measure()
                ch_points.append({"level": level, "Lv": mres.Lv, "result": mres})
                callback("gamma", {"channel": ch, "level": level, "result": mres})

            # Compute per-point gamma
            levels = [p["level"] for p in ch_points]
            lvs    = [p["Lv"]    for p in ch_points]
            gammas = calc_gamma(levels, lvs)
            for i, g_val in enumerate(gammas):
                ch_points[i]["gamma"] = g_val

            gamma_data[ch] = ch_points

        # ── 2. Color patches ────────────────────────────────────────────
        color_results: Dict[str, MeasureResult] = {}

        for name, r, g, b in (_COLOR_PATCHES if run_colors else []):
            if self._stop_requested:
                break
            cfg = PatternConfig(
                type="full_field", color=name,
                r=r, g=g, b=b,
            )
            gen.set_pattern(cfg)
            with self._engine.meter_lock:
                mres = meter.measure()
            color_results[name] = mres
            callback("color", {"name": name, "result": mres})

        # ── 3. Gamut overlap ────────────────────────────────────────────
        gamut_stats: Dict[str, float] = {}
        r_r = color_results.get("R")
        r_g = color_results.get("G")
        r_b = color_results.get("B")
        if r_r and r_g and r_b:
            gamut_stats = calc_gamut_stats(
                (r_r.u_prime, r_r.v_prime),
                (r_g.u_prime, r_g.v_prime),
                (r_b.u_prime, r_b.v_prime),
            )

        # ── 4. Δu'v' per patch ─────────────────────────────────────────
        delta_uv: Dict[str, float] = {}
        for name, mres in color_results.items():
            if name in ref_uv:
                ru, rv = ref_uv[name]
                du = mres.u_prime - ru
                dv = mres.v_prime - rv
                delta_uv[name] = math.sqrt(du * du + dv * dv) * 10_000.0

        result = {
            "gamma":       gamma_data,
            "colors":      color_results,
            "gamut_stats": gamut_stats,
            "delta_uv":    delta_uv,
        }
        callback("done", result)
        return result
