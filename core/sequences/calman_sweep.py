"""Calman-style colour saturation sweep.

Patches: R, G, B, C, M, Y  ×  20 / 40 / 60 / 80 / 100 % saturation
Total  : 30 full-field measurements

Pattern generation (Calman style):
  - Each patch = linear interpolation from White(255,255,255) to primary colour
  - e.g. Red 40% → RGB = (255, 153, 153)  [white + 40% toward full red]

Reference target calculation (same linear XYZ mixing):
  - tXYZ = (1 - w) × XYZ_white_measured  +  w × XYZ_primary100_measured
  - where w = sat_pct / 100
  - White and primaries come from ModuleMeasureSequence color patches

ΔE: CIE 1976 (ΔEab) — Euclidean distance in L*a*b* with measured white point
"""
from __future__ import annotations

import math
from typing import TYPE_CHECKING, Callable, Dict, List, Optional, Tuple

from ..equipment.base import MeasureResult, PatternConfig

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# ---------------------------------------------------------------------------
# Patch definition
# ---------------------------------------------------------------------------

SAT_LEVELS: List[int] = [20, 40, 60, 80, 100]
COLOR_ORDER: List[str] = ["R", "G", "B", "C", "M", "Y"]
GAMUT_NAMES: List[str] = ["BT.709", "DCI-P3", "BT.2020"]   # kept for UI compat

# 100% primary RGB codes (full-field, no white blend)
_PRIMARY_RGB: Dict[str, Tuple[int, int, int]] = {
    "R": (255,   0,   0),
    "G": (  0, 255,   0),
    "B": (  0,   0, 255),
    "C": (  0, 255, 255),
    "M": (255,   0, 255),
    "Y": (255, 255,   0),
}


def sat_rgb(color: str, sat_pct: int) -> Tuple[int, int, int]:
    """Return (R, G, B) 0-255 for a Calman-style saturation patch.

    Pattern = linear interpolation from White(255,255,255) toward the
    100% primary, exactly as Portrait Displays Calman sends to the display.
    """
    pr, pg, pb = _PRIMARY_RGB[color]
    w = sat_pct / 100.0
    return (
        round(255 * (1 - w) + pr * w),
        round(255 * (1 - w) + pg * w),
        round(255 * (1 - w) + pb * w),
    )


# ---------------------------------------------------------------------------
# XYZ helpers
# ---------------------------------------------------------------------------

def _xyY_to_XYZ(x: float, y: float, Y: float) -> Tuple[float, float, float]:
    if y < 1e-9:
        return 0.0, 0.0, 0.0
    return x * Y / y, Y, (1.0 - x - y) * Y / y


def meas_to_xyz(r: MeasureResult) -> Tuple[float, float, float]:
    return _xyY_to_XYZ(r.x, r.y, r.Lv)


# ---------------------------------------------------------------------------
# XYZ → Lab with custom white point
# ---------------------------------------------------------------------------

def _f(t: float) -> float:
    d = 6.0 / 29.0
    return t ** (1.0 / 3.0) if t > d ** 3 else t / (3.0 * d * d) + 4.0 / 29.0


def xyz_to_lab(X: float, Y: float, Z: float,
               Xn: float, Yn: float, Zn: float) -> Tuple[float, float, float]:
    Xn = Xn if Xn > 1e-9 else 1e-9
    Yn = Yn if Yn > 1e-9 else 1e-9
    Zn = Zn if Zn > 1e-9 else 1e-9
    fx, fy, fz = _f(X / Xn), _f(Y / Yn), _f(Z / Zn)
    return 116.0 * fy - 16.0, 500.0 * (fx - fy), 200.0 * (fy - fz)


# ---------------------------------------------------------------------------
# ΔE76 (CIE 1976 / ΔEab)
# ---------------------------------------------------------------------------

def de76(L1: float, a1: float, b1: float,
         L2: float, a2: float, b2: float) -> float:
    return math.sqrt((L1 - L2) ** 2 + (a1 - a2) ** 2 + (b1 - b2) ** 2)


# ---------------------------------------------------------------------------
# Linear XYZ reference target (Calman method)
# ---------------------------------------------------------------------------

# Secondary colours are derived from additive XYZ of their component primaries,
# matching Calman's reference model (not separately measured C/M/Y patches).
_SECONDARY_PRIMARIES: Dict[str, Tuple[str, str]] = {
    "C": ("G", "B"),
    "M": ("R", "B"),
    "Y": ("R", "G"),
}


def _primary_xyz(
    color: str,
    measured_colors: Dict[str, MeasureResult],
) -> Optional[Tuple[float, float, float]]:
    """Return the 100 % primary XYZ for a colour.

    Primaries (R/G/B/W): use measured value directly.
    Secondaries (C/M/Y): additive XYZ sum of component primaries,
    matching Calman's reference derivation method.
    """
    if color in _SECONDARY_PRIMARIES:
        p1_name, p2_name = _SECONDARY_PRIMARIES[color]
        p1 = measured_colors.get(p1_name)
        p2 = measured_colors.get(p2_name)
        if p1 is None or p2 is None:
            return None
        X1, Y1, Z1 = meas_to_xyz(p1)
        X2, Y2, Z2 = meas_to_xyz(p2)
        return X1 + X2, Y1 + Y2, Z1 + Z2
    else:
        p = measured_colors.get(color)
        if p is None:
            return None
        return meas_to_xyz(p)


def calc_target_xyz(
    color: str,
    sat_pct: int,
    measured_colors: Dict[str, MeasureResult],
) -> Optional[Tuple[float, float, float]]:
    """Compute the reference XYZ for one patch via linear XYZ mixing.

    target = (1-w) × XYZ_white  +  w × XYZ_primary100
    where w = sat_pct / 100

    For secondary colours (C/M/Y), XYZ_primary100 is derived from the additive
    sum of component measured primaries (Calman method).
    """
    w_meas = measured_colors.get("W")
    if w_meas is None:
        return None
    primary_xyz = _primary_xyz(color, measured_colors)
    if primary_xyz is None:
        return None
    w = sat_pct / 100.0
    Xn, Yn, Zn = meas_to_xyz(w_meas)
    Xp, Yp, Zp = primary_xyz
    return (
        (1 - w) * Xn + w * Xp,
        (1 - w) * Yn + w * Yp,
        (1 - w) * Zn + w * Zp,
    )


def calc_de76(
    color: str,
    sat_pct: int,
    meas: MeasureResult,
    measured_colors: Dict[str, MeasureResult],
) -> float:
    """ΔE76 between measured patch and Calman-style linear XYZ reference."""
    w_meas = measured_colors.get("W")
    if w_meas is None:
        return 0.0
    Xn, Yn, Zn = meas_to_xyz(w_meas)
    white_xyz = (Xn, Yn, Zn)

    ref_xyz = calc_target_xyz(color, sat_pct, measured_colors)
    if ref_xyz is None:
        return 0.0

    Xr, Yr, Zr = ref_xyz
    L_ref, a_ref, b_ref = xyz_to_lab(Xr, Yr, Zr, *white_xyz)

    Xm, Ym, Zm = meas_to_xyz(meas)
    L_meas, a_meas, b_meas = xyz_to_lab(Xm, Ym, Zm, *white_xyz)

    return de76(L_ref, a_ref, b_ref, L_meas, a_meas, b_meas)


# ---------------------------------------------------------------------------
# Compatibility shim — kept so module_panel._add_row() still imports cleanly
# ---------------------------------------------------------------------------

def extract_gamut_data(_measured_colors: Dict[str, MeasureResult]):
    """Kept for UI import compatibility; no longer used for reference calc."""
    return None


# ---------------------------------------------------------------------------
# Sequence
# ---------------------------------------------------------------------------

class CalmanSweepSequence:
    """Calman-style saturation sweep using measured display primaries.

    Requires measured_colors (from ModuleMeasureSequence) with at least
    R, G, B, C, M, Y, W entries.
    """

    def __init__(self, engine: "MeasurementEngine") -> None:
        self._engine = engine
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def run(
        self,
        is_hdr: bool,
        measured_colors: Dict[str, MeasureResult],
        callback: Callable[[str, int, MeasureResult, float], None],
    ) -> Dict[str, Dict[int, Dict]]:
        """Run the saturation sweep.

        Parameters
        ----------
        is_hdr:           HDR output mode.
        measured_colors:  Color patch results from ModuleMeasureSequence.
                          Needs R, G, B, C, M, Y, W for reference targets.
        callback:         Called per patch: (color, sat_pct, result, de76).

        Returns
        -------
        {color: {sat_pct: {"result": MeasureResult, "de76": float,
                           "ref_lab": (L,a,b), "meas_lab": (L,a,b)}}}
        """
        if "W" not in measured_colors:
            raise ValueError(
                "White(W) 측정 결과가 필요합니다. "
                "먼저 '모듈 전체' 측정을 실행하세요."
            )

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

        Xn, Yn, Zn = meas_to_xyz(measured_colors["W"])
        white_xyz = (Xn, Yn, Zn)

        self._stop_requested = False
        results: Dict[str, Dict[int, Dict]] = {}

        for color in COLOR_ORDER:
            if self._stop_requested:
                break
            color_data: Dict[int, Dict] = {}

            for sat in SAT_LEVELS:
                if self._stop_requested:
                    break

                r, g, b = sat_rgb(color, sat)
                cfg = PatternConfig(
                    type="full_field",
                    color=f"{color}_{sat}",
                    r=r, g=g, b=b,
                )
                gen.set_pattern(cfg)
                with self._engine.meter_lock:
                    mres = meter.measure()

                # Reference: linear XYZ mix of measured white and primary
                ref_xyz = calc_target_xyz(color, sat, measured_colors)
                if ref_xyz is not None:
                    Xr, Yr, Zr = ref_xyz
                    ref_lab = xyz_to_lab(Xr, Yr, Zr, *white_xyz)
                else:
                    ref_lab = (0.0, 0.0, 0.0)

                Xm, Ym, Zm = meas_to_xyz(mres)
                meas_lab = xyz_to_lab(Xm, Ym, Zm, *white_xyz)
                delta = de76(*ref_lab, *meas_lab)

                color_data[sat] = {
                    "result":   mres,
                    "de76":     delta,
                    "ref_lab":  ref_lab,
                    "meas_lab": meas_lab,
                }
                callback(color, sat, mres, delta)

            results[color] = color_data

        return results
