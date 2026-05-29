from __future__ import annotations

import threading
import time
from typing import TYPE_CHECKING, Callable, List

from ..equipment.base import MeasureResult, PatternConfig, PatternInfo

if TYPE_CHECKING:
    from ..engine import MeasurementEngine

# 10% APL 정사각형 윈도우: W=31.6%, H=31.6% (블랙 바탕 + 흰색 박스)
# 31.6² / 100 = 9.9856% ≈ 10% APL
_APL_PCT = 10.0
_WINDOW_SIDE_PCT = 31.6  # % (H=31.6, W=31.6 — 규격 고정값)


class LumSwingSequence:
    """Continuous luminance swing measurement — 301 samples at 1 sample/sec (~5 min).

    Streams MeasureResult objects to callback in real-time so the UI can plot
    a live time-vs-Lv chart while measurement is in progress.
    """

    DEFAULT_SAMPLE_COUNT = 301
    DEFAULT_INTERVAL_SEC = 1.0  # 1 sample per second

    def __init__(self, engine: "MeasurementEngine") -> None:
        self._engine = engine
        self._stop_event = threading.Event()

    def stop(self) -> None:
        """Signal the measurement loop to stop after the current sample."""
        self._stop_event.set()

    def run(
        self,
        case: str,
        is_hdr: bool,
        callback: Callable[[MeasureResult], None],
        sample_count: int = DEFAULT_SAMPLE_COUNT,
        interval_sec: float = DEFAULT_INTERVAL_SEC,
    ) -> List[MeasureResult]:
        """Run the luminance swing sequence for one case (A/B/C).

        Parameters
        ----------
        case:          Label for this measurement run ("A", "B", or "C").
        is_hdr:        If True the generator is switched to HDR before outputting.
        callback:      Called with each MeasureResult as it arrives (real-time).
        sample_count:  Total number of samples (default 301).
        interval_sec:  Target interval between samples in seconds (default 1.0).
                       Measurement time is counted inside the interval; only the
                       remainder is slept so the stop signal can interrupt early.

        Returns the complete list of MeasureResult collected.
        """
        gen = self._engine.generator
        meter = self._engine.meter
        if gen is None or not gen.is_connected:
            raise RuntimeError("Generator is not connected")
        if meter is None or not meter.is_connected:
            raise RuntimeError("Meter is not connected")

        self._stop_event.clear()

        # Switch HDR/SDR before pattern output
        if is_hdr:
            gen.set_hdr(True)
        else:
            gen.set_sdr()

        pattern_info = PatternInfo(
            type="window",
            apl_pct=10.0,
            width_pct=_WINDOW_SIDE_PCT,
            height_pct=_WINDOW_SIDE_PCT,
            color="white",
            is_hdr=is_hdr,
        )

        cfg = PatternConfig(
            type="window",
            color="white",
            r=255, g=255, b=255,
            width_pct=_WINDOW_SIDE_PCT,
            height_pct=_WINDOW_SIDE_PCT,
            bg_r=0, bg_g=0, bg_b=0,
        )
        gen.set_pattern(cfg)

        # Stamp the current pattern on the meter so each MeasureResult carries it
        if hasattr(meter, "set_current_pattern"):
            meter.set_current_pattern(pattern_info)  # type: ignore[attr-defined]

        results: List[MeasureResult] = []
        for _ in range(sample_count):
            if self._stop_event.is_set():
                break
            t0 = time.monotonic()
            with self._engine.meter_lock:
                result = meter.measure()
            results.append(result)
            callback(result)
            # Sleep the remainder of the interval; interruptible by stop()
            remaining = interval_sec - (time.monotonic() - t0)
            if remaining > 0:
                self._stop_event.wait(timeout=remaining)

        return results
