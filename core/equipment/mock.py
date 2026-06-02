"""Mock meter and generator for offline testing / demos."""
from __future__ import annotations

import math
import random
import time

from .base import GeneratorBase, MeasureResult, MeterBase, PatternConfig, PatternInfo


class MockMeter(MeterBase):
    """Simulated colorimeter that returns plausible random values."""

    def __init__(self) -> None:
        self._connected = True
        self._pattern: PatternInfo | None = None

    def connect(self, port: str) -> None:
        self._connected = True

    def disconnect(self) -> None:
        self._connected = False

    def set_current_pattern(self, info: PatternInfo) -> None:
        self._pattern = info

    def measure(self) -> MeasureResult:
        time.sleep(0.05)  # simulate measurement delay
        p = self._pattern

        # Base luminance depends on pattern colour / APL
        if p is None:
            base_lv = 300.0
        elif p.color == "white":
            base_lv = 300.0 * (p.apl_pct / 100.0) ** 0.5 if p.apl_pct else 300.0
        elif p.color == "red":
            base_lv = 90.0
        elif p.color == "green":
            base_lv = 180.0
        elif p.color == "blue":
            base_lv = 30.0
        elif p.color == "black":
            base_lv = 0.01
        else:
            base_lv = 100.0

        if p and p.type in ("raster_window", "window") and p.color == "black":
            # Black window — very low luminance
            base_lv = random.uniform(0.008, 0.015)
        else:
            base_lv *= random.uniform(0.97, 1.03)

        # Chromaticity
        if p and p.color == "red":
            x, y = 0.680 + random.gauss(0, 0.002), 0.320 + random.gauss(0, 0.002)
        elif p and p.color == "green":
            x, y = 0.265 + random.gauss(0, 0.002), 0.690 + random.gauss(0, 0.002)
        elif p and p.color == "blue":
            x, y = 0.150 + random.gauss(0, 0.002), 0.060 + random.gauss(0, 0.002)
        elif p and p.color == "black":
            x, y = 0.313 + random.gauss(0, 0.005), 0.329 + random.gauss(0, 0.005)
        else:
            x, y = 0.313 + random.gauss(0, 0.002), 0.329 + random.gauss(0, 0.002)

        denom = -2 * x + 12 * y + 3
        u_prime = 4 * x / denom if denom else 0.0
        v_prime = 9 * y / denom if denom else 0.0

        Y = base_lv
        X = Y * x / y if y else 0.0
        Z = Y * (1 - x - y) / y if y else 0.0

        cct = 6500.0 + random.gauss(0, 50)
        duv = random.gauss(0.0, 0.002)

        return MeasureResult(
            timestamp_ms=int(time.time() * 1000),
            Lv=round(base_lv, 4),
            x=round(x, 4),
            y=round(y, 4),
            u_prime=round(u_prime, 4),
            v_prime=round(v_prime, 4),
            X=round(X, 4),
            Y=round(Y, 4),
            Z=round(Z, 4),
            cct=round(cct, 1),
            duv=round(duv, 5),
            pattern_info=p or PatternInfo(
                type="unknown", apl_pct=0.0, width_pct=0.0, height_pct=0.0, color="unknown"
            ),
        )

    @property
    def is_connected(self) -> bool:
        return self._connected


class MockGenerator(GeneratorBase):
    """Simulated pattern generator that does nothing but track state."""

    def __init__(self) -> None:
        self._connected = True
        self._hdr = False
        self._current_pattern: PatternConfig | None = None

    def connect(self, port: str) -> None:
        self._connected = True

    def disconnect(self) -> None:
        self._connected = False

    def set_pattern(self, cfg: PatternConfig) -> None:
        self._current_pattern = cfg

    def set_hdr(self, enabled: bool) -> None:
        self._hdr = enabled

    def set_sdr(self) -> None:
        self._hdr = False

    def show_black(self) -> None:
        pass

    @property
    def is_connected(self) -> bool:
        return self._connected
