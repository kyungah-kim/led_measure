from __future__ import annotations

import time
from typing import TYPE_CHECKING, Callable, Dict, List

from ..equipment.base import MeasureResult, PatternConfig, PatternInfo

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# APL step sets (percent values) — all ordered 100 → smallest (descending pattern size)
STEPS_37: List[int] = [
    100, 95, 90, 86, 81, 77, 72, 68, 64, 60, 56, 53, 49, 46, 42,
    39, 36, 33, 30, 28, 25, 23, 20, 18, 16, 14, 12, 11, 10, 9, 8,
    6, 5, 4, 3, 2, 1,
]

STEPS_11: List[int] = [100, 81, 64, 49, 36, 25, 16, 10, 4, 3, 1]

STEPS_10: List[int] = [100, 81, 64, 49, 36, 25, 16, 10, 3, 1]

STEPS_2: List[int] = [100, 10]

_STEP_VERSIONS: Dict[str, List[int]] = {
    "37": STEPS_37,
    "11": STEPS_11,
    "10": STEPS_10,
    "2": STEPS_2,
}

# APL threshold below which cooling is applicable
_COOLING_APL_THRESHOLD = 10

# Duration of the black cooling screen in seconds
_COOLING_DURATION_SEC = 5.0

# Default number of measurements per APL step
_MEASUREMENTS_PER_STEP = 3

# Interval between repeated measurements at the same APL (seconds).
# Display is already stable — only need meter to clear its buffer between readings.
_INTER_MEAS_SLEEP = 0.3


def _apl_to_window_size(apl_pct: float) -> tuple[float, float]:
    """Convert APL % to symmetric window dimensions (W%, H%).

    For a square-centred window: APL = (W/100) * (H/100) * 100
    So W=H=sqrt(APL/100)*100.
    """
    side = (apl_pct / 100.0) ** 0.5 * 100.0
    return round(side, 2), round(side, 2)


class LumLoadingSequence:
    """Luminance Loading (APL sweep) sequence.

    Supports 37-step, 10-step, and 2-step APL versions.
    Runs N measurements per APL step with optional black-screen cooling
    for APL values ≤ 10.
    """

    def __init__(self, engine: "MeasurementEngine") -> None:
        self._engine = engine
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def run(
        self,
        version: str,
        case: str,  # noqa: ARG002  (caller uses for labelling; sequence ignores)
        is_hdr: bool,
        cooling_enabled: bool,
        callback: Callable[[int, float, List[MeasureResult]], None],
        measurements_per_step: int = _MEASUREMENTS_PER_STEP,
        cooling_duration_sec: float = _COOLING_DURATION_SEC,
        cooling_apl_threshold: int = _COOLING_APL_THRESHOLD,
    ) -> Dict[int, List[MeasureResult]]:
        """Execute the APL sweep for one case.

        Parameters
        ----------
        version:              "37", "10", or "2".
        case:                 Case label ("Vivid", "Standard", "Cinema").
        is_hdr:               Apply HDR settings before the sweep.
        cooling_enabled:      Insert black cooling screen before low-APL steps.
        callback:             Called after each step with (step_index, apl_pct, results).
        measurements_per_step: Number of measurements to take per APL step.
        cooling_duration_sec: How long to show black screen when cooling.

        Returns dict mapping APL% -> list of MeasureResult.
        """
        steps = _STEP_VERSIONS.get(version)
        if steps is None:
            raise ValueError(f"Unknown version {version!r}. Choose '37', '10', or '2'.")

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
        results: Dict[int, List[MeasureResult]] = {}

        for step_idx, apl in enumerate(steps):
            if self._stop_requested:
                break
            w_pct, h_pct = _apl_to_window_size(float(apl))

            # Optional cooling: 직통 블랙 (ALLCLR4+EXPDN4) — 타이밍 로드 없이 즉시 검정
            if cooling_enabled and apl <= cooling_apl_threshold:
                if hasattr(gen, "show_black"):
                    gen.show_black()  # type: ignore[attr-defined]
                time.sleep(cooling_duration_sec)

            # Output the target APL window pattern
            pattern_info = PatternInfo(
                type="window",
                apl_pct=float(apl),
                width_pct=w_pct,
                height_pct=h_pct,
                color="white",
                is_hdr=is_hdr,
            )
            cfg = PatternConfig(
                type="window",
                color="white",
                r=255, g=255, b=255,
                width_pct=w_pct,
                height_pct=h_pct,
                bg_r=0, bg_g=0, bg_b=0,
            )
            gen.set_pattern(cfg)

            if hasattr(meter, "set_current_pattern"):
                meter.set_current_pattern(pattern_info)  # type: ignore[attr-defined]

            step_results: List[MeasureResult] = []
            for i in range(measurements_per_step):
                if i > 0:
                    time.sleep(_INTER_MEAS_SLEEP)
                with self._engine.meter_lock:
                    result = meter.measure()
                step_results.append(result)

            results[apl] = step_results
            callback(step_idx, float(apl), step_results)

        return results
