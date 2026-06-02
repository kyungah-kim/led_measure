from __future__ import annotations

from typing import TYPE_CHECKING, Callable, Dict, List

from ..equipment.base import MeasureResult, PatternConfig, PatternInfo

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# 검은 창 H/V 크기 스텝 (% of screen side).
# White Raster 위에 중앙 검은 창을 이 크기로 순서대로 출력한다.
_WIN_SIDES_PCT: List[float] = [100.0, 50.0, 20.0, 14.1]


class ContrastSequence:
    """Contrast Ratio measurement — White Raster + centered Black Window.

    패턴: SPTS4(0,1,2,10) White Raster + 중앙 Black Window.
      창 100% → 전체 화면 검은 창 (최대 부하 상태 기준값)
      창  50% → 50%×50% 검은 창
      창  20% → 20%×20% 검은 창
      창 14.1% → 14.1%×14.1% 검은 창

    측정기는 항상 화면 중앙(검은 창 내부)을 향한다.
    매 스텝 White Raster를 재출력한 뒤 검은 창 크기를 변경한다.
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
        """Measure full white first, then luminance at each black window size.

        Parameters
        ----------
        is_hdr:   Apply HDR output settings.
        callback: Called with (win_side_pct, MeasureResult) after each step.
                  win_side_pct == 0.0 → full white reference measurement.

        Returns dict mapping win_side_pct -> MeasureResult.
          0.0      → full white (reference)
          100/50/20/14.1 → black window sizes (CR = white_lv / window_lv)
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

        # ── Step 0: Full White reference ──────────────────────────────────────
        white_info = PatternInfo(
            type="full_field",
            apl_pct=100.0,
            width_pct=100.0,
            height_pct=100.0,
            color="white",
            is_hdr=is_hdr,
        )
        white_cfg = PatternConfig(
            type="full_field",
            color="white",
            r=255, g=255, b=255,
            width_pct=100.0,
            height_pct=100.0,
        )
        gen.set_pattern(white_cfg)
        if hasattr(meter, "set_current_pattern"):
            meter.set_current_pattern(white_info)  # type: ignore[attr-defined]
        with self._engine.meter_lock:
            white_result = meter.measure()
        results[0.0] = white_result
        callback(0.0, white_result)

        # ── Steps 1‥N: Black window sizes ────────────────────────────────────
        for side_pct in _WIN_SIDES_PCT:
            if self._stop_requested:
                break

            pattern_info = PatternInfo(
                type="raster_window",
                apl_pct=0.0,
                width_pct=side_pct,
                height_pct=side_pct,
                color="black",
                is_hdr=is_hdr,
            )
            cfg = PatternConfig(
                type="raster_window",
                color="black",
                r=0, g=0, b=0,
                width_pct=side_pct,
                height_pct=side_pct,
            )
            gen.set_pattern(cfg)

            if hasattr(meter, "set_current_pattern"):
                meter.set_current_pattern(pattern_info)  # type: ignore[attr-defined]

            with self._engine.meter_lock:
                result = meter.measure()

            results[side_pct] = result
            callback(side_pct, result)

        return results
