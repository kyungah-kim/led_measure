from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Dict, List

from ..equipment.base import MeasureResult, PatternConfig, PatternInfo

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# Gamut measurement order and full-field RGB values
_GAMUT_SEQUENCE: List[tuple[str, int, int, int]] = [
    ("red",   255,   0,   0),
    ("green",   0, 255,   0),
    ("blue",    0,   0, 255),
    ("white", 255, 255, 255),
    ("black",   0,   0,   0),
]


class GamutSequence:
    """Full-screen colour gamut measurement.

    Outputs each primary colour (R→G→B→W→BK) as a 100% full-field pattern
    and takes one measurement per colour.
    """

    def __init__(self, engine: "MeasurementEngine") -> None:
        self._engine = engine
        self._stop_requested = False

    def stop(self) -> None:
        self._stop_requested = True

    def run(
        self,
        is_hdr: bool,
        callback: Callable[[str, MeasureResult], None],
    ) -> Dict[str, MeasureResult]:
        """Measure R, G, B, White, Black in sequence.

        Parameters
        ----------
        is_hdr:   Apply HDR output settings.
        callback: Called with (color_name, MeasureResult) after each measurement.

        Returns dict mapping colour name -> MeasureResult.
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

        results: Dict[str, MeasureResult] = {}

        self._stop_requested = False

        for color_name, r, g, b in _GAMUT_SEQUENCE:
            if self._stop_requested:
                break
            pattern_info = PatternInfo(
                type="full_field",
                apl_pct=100.0 if color_name != "black" else 0.0,
                width_pct=100.0,
                height_pct=100.0,
                color=color_name,
                is_hdr=is_hdr,
            )
            cfg = PatternConfig(type="full_field", color=color_name, r=r, g=g, b=b)
            gen.set_pattern(cfg)

            if hasattr(meter, "set_current_pattern"):
                meter.set_current_pattern(pattern_info)  # type: ignore[attr-defined]

            with self._engine.meter_lock:
                result = meter.measure()

            results[color_name] = result
            callback(color_name, result)

        return results
