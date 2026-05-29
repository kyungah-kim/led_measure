from __future__ import annotations

import time
from typing import TYPE_CHECKING, Callable, Dict, List

from ..equipment.base import MeasureResult, PatternConfig, PatternInfo

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# APL step sets (percent values)
STEPS_37: List[int] = [
    1, 2, 3, 4, 5, 6, 8, 9, 10, 11, 12, 14, 16, 18, 20,
    23, 25, 28, 30, 33, 36, 39, 42, 46, 49, 53, 56, 60,
    64, 68, 72, 77, 81, 86, 90, 95, 100,
]

STEPS_10: List[int] = [1, 3, 10, 16, 25, 36, 49, 64, 81, 100]

STEPS_2: List[int] = [10, 100]

_STEP_VERSIONS: Dict[str, List[int]] = {
    "37": STEPS_37,
    "10": STEPS_10,
    "2": STEPS_2,
}

# APL threshold below which cooling is applicable
_COOLING_APL_THRESHOLD = 10

# Duration of the black cooling screen in seconds
_COOLING_DURATION_SEC = 5.0

# Default number of measurements per APL step
_MEASUREMENTS_PER_STEP = 5


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
        case: str,
        is_hdr: bool,
        cooling_enabled: bool,
        callback: Callable[[int, float, List[MeasureResult]], None],
        measurements_per_step: int = _MEASUREMENTS_PER_STEP,
        cooling_duration_sec: float = _COOLING_DURATION_SEC,
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
        steps = _STEP_VERSIONS.get(str(version))
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

            # Optional cooling: show black before low-APL steps
            if cooling_enabled and apl <= _COOLING_APL_THRESHOLD:
                black_cfg = PatternConfig(
                    type="full_field", color="black", r=0, g=0, b=0
                )
                gen.set_pattern(black_cfg)
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
            for _ in range(measurements_per_step):
                with self._engine.meter_lock:
                    result = meter.measure()
                step_results.append(result)

            results[apl] = step_results
            callback(step_idx, float(apl), step_results)

        return results
