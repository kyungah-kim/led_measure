from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Dict, List

from ..equipment.base import MeasureResult, PatternConfig, PatternInfo

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# APL values to sweep.  Each value is a percentage of total screen area.
# The window linear side is sqrt(apl/100)*100 so the actual lit area equals apl%.
_APL_SIZES_PCT: List[float] = [100.0, 50.0, 20.0, 14.1, 0.0]


def _apl_to_window_pct(apl: float) -> float:
    """Return the linear side length (% of screen dimension) for a given APL%."""
    return (apl / 100.0) ** 0.5 * 100.0


class ContrastSequence:
    """Contrast Ratio measurement across multiple window sizes.

    Background: White raster (100%)
    Foreground window: Black, sweeping through 100/50/20/14.1/0% sizes.
    A 0% window effectively gives a pure white measurement.
    """

    def __init__(self, engine: "MeasurementEngine") -> None:
        self._engine = engine
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def run(
        self,
        is_hdr: bool,
        callback: Callable[[float, MeasureResult], None],
    ) -> Dict[float, MeasureResult]:
        """Measure luminance at each window size.

        Parameters
        ----------
        is_hdr:   Apply HDR output settings.
        callback: Called with (window_size_pct, MeasureResult) after each step.

        Returns dict mapping window_size_pct -> MeasureResult.
        """
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

        results: Dict[float, MeasureResult] = {}

        self._stop_requested = False

        for apl in _APL_SIZES_PCT:
            if self._stop_requested:
                break
            side_pct = _apl_to_window_pct(apl)
            if apl == 0.0:
                # 0% window = solid white raster — meter reads peak white
                pattern_info = PatternInfo(
                    type="full_field",
                    apl_pct=100.0,
                    width_pct=100.0,
                    height_pct=100.0,
                    color="white",
                    is_hdr=is_hdr,
                )
                cfg = PatternConfig(
                    type="full_field", color="white", r=255, g=255, b=255
                )
            else:
                # White background + centred black window — meter aimed at black window.
                # side_pct is the linear dimension so actual black area = apl%.
                pattern_info = PatternInfo(
                    type="window",
                    apl_pct=0.0,          # meter sees black center → near-zero Lv
                    width_pct=side_pct,
                    height_pct=side_pct,
                    color="black",
                    is_hdr=is_hdr,
                )
                cfg = PatternConfig(
                    type="window",
                    color="black",
                    r=0, g=0, b=0,
                    width_pct=side_pct,
                    height_pct=side_pct,
                    bg_r=255, bg_g=255, bg_b=255,
                )
            gen.set_pattern(cfg)

            if hasattr(meter, "set_current_pattern"):
                meter.set_current_pattern(pattern_info)  # type: ignore[attr-defined]

            with self._engine.meter_lock:
                result = meter.measure()

            results[apl] = result
            callback(apl, result)

        return results
