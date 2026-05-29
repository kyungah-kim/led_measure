"""Mock meter and generator for UI/output testing without physical hardware."""
from __future__ import annotations

import random
import time

from .base import GeneratorBase, MeasureResult, MeterBase, PatternConfig, PatternInfo
from ..colorimetry import xy_to_cct_duv

# Realistic display primary chromaticities (sRGB-like)
_COLOR_XY: dict[str, tuple[float, float]] = {
    "red":     (0.6400, 0.3300),
    "green":   (0.3000, 0.6000),
    "blue":    (0.1500, 0.0600),
    "white":   (0.3127, 0.3290),
    "black":   (0.3127, 0.3290),
    "cyan":    (0.2250, 0.3295),
    "magenta": (0.3950, 0.1650),
    "yellow":  (0.4190, 0.5050),
}

# Peak luminance per colour at 100% APL (cd/m²)
_COLOR_LV_PEAK: dict[str, float] = {
    "red":     108.0,
    "green":   367.0,
    "blue":     35.0,
    "white":   500.0,
    "black":     0.008,
    "cyan":    380.0,
    "magenta": 130.0,
    "yellow":  450.0,
}


def _xy_to_uv(x: float, y: float) -> tuple[float, float]:
    d = -2 * x + 12 * y + 3
    if d == 0:
        return 0.0, 0.0
    return 4 * x / d, 9 * y / d


def _noise(val: float, pct: float = 0.003) -> float:
    return val + random.gauss(0, abs(val) * pct + 1e-9)


class MockMeter(MeterBase):
    """Simulates CA-410 with realistic colour-dependent values.

    Chromaticity changes per pattern colour; luminance scales with APL.
    """

    def __init__(self) -> None:
        self._connected = False
        self._current_pattern = PatternInfo(
            type="window", apl_pct=10.0,
            width_pct=31.6, height_pct=31.6, color="white"
        )

    def connect(self, port: str) -> None:
        time.sleep(0.05)
        self._connected = True

    def disconnect(self) -> None:
        self._connected = False

    @property
    def is_connected(self) -> bool:
        return self._connected

    def set_current_pattern(self, info: PatternInfo) -> None:
        self._current_pattern = info

    def measure(self) -> MeasureResult:
        """Return a realistic measurement based on the current pattern.

        CA-410 integration time:
          - Lv < 0.01  cd/m²  → ~3.0 s (very dark, long integration)
          - Lv < 1.0   cd/m²  → ~2.0 s (dark, extended integration)
          - Lv >= 1.0  cd/m²  → ~1.5 s (normal auto-range)
        """
        p = self._current_pattern
        color = p.color.lower()

        # Base chromaticity for this colour
        base_x, base_y = _COLOR_XY.get(color, _COLOR_XY["white"])

        # Add small chromaticity noise
        x = max(0.01, min(0.85, _noise(base_x, 0.002)))
        y = max(0.01, min(0.85, _noise(base_y, 0.002)))

        # Luminance: peak × APL factor (with slight HDR boost)
        peak = _COLOR_LV_PEAK.get(color, 500.0)
        if p.is_hdr:
            peak *= 2.0   # simple HDR luminance boost

        apl_factor = p.apl_pct / 100.0
        # Non-linear (gamma-like) APL-to-Lv relationship
        lv_base = peak * (apl_factor ** 0.45) if color != "black" else peak
        Lv = max(0.001, _noise(lv_base, 0.005))

        # CA-410 integration time depends on luminance level
        if Lv < 0.01:
            time.sleep(3.0)
        elif Lv < 1.0:
            time.sleep(2.0)
        else:
            time.sleep(1.5)

        # Derive XYZ from xyY
        Y = Lv
        X = (x / y) * Y if y > 0 else 0.0
        Z = ((1 - x - y) / y) * Y if y > 0 else 0.0

        u_prime, v_prime = _xy_to_uv(x, y)
        cct, duv = xy_to_cct_duv(x, y)

        return MeasureResult(
            timestamp_ms=int(time.time() * 1000),
            Lv=round(Lv, 4),
            x=round(x, 4),
            y=round(y, 4),
            u_prime=round(u_prime, 4),
            v_prime=round(v_prime, 4),
            X=round(X, 4),
            Y=round(Y, 4),
            Z=round(Z, 4),
            cct=cct,
            duv=duv,
            pattern_info=self._current_pattern,
        )


class MockGenerator(GeneratorBase):
    """Simulates VG-876/VG-879 without actual hardware output.

    Pattern stabilization delay (2 s) models the time for the TV panel
    to respond to a new signal from the generator before the meter can
    take a valid reading.
    """

    _STABILIZE_SEC = 2.0   # display stabilization after pattern change

    def __init__(self) -> None:
        self._connected = False
        self._hdr = False
        self._current: PatternConfig | None = None

    def connect(self, port: str) -> None:
        time.sleep(0.02)
        self._connected = True

    def disconnect(self) -> None:
        self._connected = False

    @property
    def is_connected(self) -> bool:
        return self._connected

    def set_pattern(self, cfg: PatternConfig) -> None:
        self._current = cfg
        time.sleep(self._STABILIZE_SEC)

    def show_full_field(self, r: int, g: int, b: int, bit_mode: int = 8) -> None:
        self._current = PatternConfig(
            type="full_field", color=f"rgb({r},{g},{b})",
            r=r, g=g, b=b, bit_mode=bit_mode,
        )
        time.sleep(self._STABILIZE_SEC)

    def show_window(self, w_pct: float, h_pct: float,
                    fg_r: int = 255, fg_g: int = 255, fg_b: int = 255,
                    bg_r: int = 0, bg_g: int = 0, bg_b: int = 0,
                    bit_mode: int = 8) -> None:
        self._current = PatternConfig(
            type="window", color="white",
            r=fg_r, g=fg_g, b=fg_b,
            width_pct=w_pct, height_pct=h_pct,
            bg_r=bg_r, bg_g=bg_g, bg_b=bg_b,
        )
        time.sleep(self._STABILIZE_SEC)

    def set_hdr(self, enabled: bool) -> None:
        self._hdr = enabled

    def set_sdr(self) -> None:
        self._hdr = False
