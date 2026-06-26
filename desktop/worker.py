from __future__ import annotations

import threading
import time
from typing import Any, Callable, Optional

from PySide6.QtCore import QThread, Signal

from core.engine import MeasurementEngine

# (mode_idx, is_hdr, case, seq_name) — 14개 개별 측정 단계
_ALL_STEPS: list[tuple[int, bool, str, str]] = [
    (0, False, "Vivid",    "lum_swing"),
    (0, False, "Vivid",    "lum_loading"),
    (0, False, "Vivid",    "gamut"),
    (0, False, "Vivid",    "contrast"),
    (1, False, "Standard", "lum_swing"),
    (1, False, "Standard", "lum_loading"),
    (2, False, "Cinema",   "lum_swing"),
    (2, False, "Cinema",   "lum_loading"),
    (3, True,  "Vivid",    "lum_swing"),
    (3, True,  "Vivid",    "lum_loading"),
    (4, True,  "Standard", "lum_swing"),
    (4, True,  "Standard", "lum_loading"),
    (5, True,  "Cinema",   "lum_swing"),
    (5, True,  "Cinema",   "lum_loading"),
]


def wire_worker_cleanup(worker: QThread, owner: object, attr: str,
                        extra_cb: Optional[Callable] = None) -> None:
    """worker.finished 에 안전한 정리 슬롯을 연결한다.

    finished 는 run() 반환 직후 발생하지만 OS 스레드가 완전히 종료되기
    전일 수 있다. wait() 를 먼저 호출해 OS 레벨 종료를 보장한 뒤
    Python 참조를 해제해야 GC 가 C++ 소멸자를 안전하게 호출할 수 있다.

    deleteLater() 를 사용하면 Python GC 와 Qt 이벤트루프의 소멸 타이밍이
    겹쳐 'QThread: Destroyed while thread is still running' 크래시가 발생.
    """
    def _cleanup() -> None:
        worker.wait()                    # OS 스레드 완전 종료 대기 (이미 끝났으면 즉시 반환)
        setattr(owner, attr, None)       # Python 참조 해제 → GC 안전 소멸
        if extra_cb is not None:
            extra_cb()

    worker.finished.connect(_cleanup)


class LgTvReadWorker(QThread):
    """LG TV 시리얼 포트에서 데이터를 읽어 signal로 전달하는 백그라운드 워커."""

    data_received = Signal(str)

    def __init__(self, serial_obj: Any) -> None:
        super().__init__()
        self._serial = serial_obj
        self._stop   = threading.Event()

    def stop(self) -> None:
        self._stop.set()

    def run(self) -> None:
        while not self._stop.is_set():
            try:
                if self._serial and self._serial.is_open and self._serial.in_waiting > 0:
                    raw = self._serial.read(self._serial.in_waiting)
                    # latin-1 은 0x00~0xFF 를 손실 없이 디코딩 — 이진 데이터 깨짐 방지
                    self.data_received.emit(raw.decode("latin-1"))
                else:
                    time.sleep(0.05)
            except Exception:
                break


class ConnectWorker(QThread):
    """Runs a blocking connect call in a background thread.

    NOTE: do NOT define a signal named 'finished' here — that would shadow
    QThread::finished (C++ signal always emitted when run() returns).
    We rely on QThread::finished for deleteLater() cleanup.
    Use 'succeeded' for the success-path notification instead.
    """

    succeeded: Signal = Signal()
    error: Signal = Signal(str)

    def __init__(self, fn: Callable[[], None]) -> None:
        super().__init__()
        self._fn = fn

    def run(self) -> None:
        try:
            self._fn()
            self.succeeded.emit()
        except Exception as exc:
            self.error.emit(str(exc))
        # QThread::finished is emitted automatically by Qt here (always),
        # triggering any connected deleteLater() regardless of success/error.


class MeasurementWorker(QThread):
    """Runs engine.run_sequence() in a background thread.

    Signals
    -------
    progress(step_name, progress_fraction, data)
        Emitted for every on_progress callback fired by the engine.
    succeeded(result)
        Emitted once the sequence returns its result dict/list.
    error(message)
        Emitted if an exception escapes the sequence.

    NOTE: 'finished' is intentionally NOT defined here so that QThread::finished
    (always emitted when run() returns) is accessible as worker.finished for
    connecting deleteLater() cleanup.
    """

    progress: Signal = Signal(str, float, object)
    succeeded: Signal = Signal(object)
    error: Signal = Signal(str)

    def __init__(
        self,
        engine: MeasurementEngine,
        seq_name: str,
        **kwargs: Any,
    ) -> None:
        super().__init__()
        self._engine = engine
        self._seq_name = seq_name
        self._kwargs = kwargs

    def run(self) -> None:
        def _progress(step: str, pct: float, data: Any) -> None:
            self.progress.emit(step, pct, data)

        original_cb = self._engine.on_progress
        self._engine.on_progress = _progress

        try:
            result = self._engine.run_sequence(self._seq_name, **self._kwargs)
            self.succeeded.emit(result)
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            self._engine.on_progress = original_cb
        # QThread::finished fires here automatically — triggers deleteLater() cleanup.


def _versioned_key(d: dict, key: str) -> str:
    """Return key if not in d, else key_v2, key_v3, … (first unused)."""
    if key not in d:
        return key
    v = 2
    while f"{key}_v{v}" in d:
        v += 1
    return f"{key}_v{v}"


class AutoAllWorker(QThread):
    """6개 PSM 모드를 순서대로 자동 측정하는 워커.

    Competitor: mode_change_requested 시 main thread 가 다이얼로그 → confirm() 호출
    LG:         mode_change_requested 시 main thread 가 시리얼 명령 전송 → confirm() 호출
    """

    mode_change_requested = Signal(bool, str)   # is_hdr, case
    step_started          = Signal(int, str)    # step_idx, label
    progress              = Signal(str, float, object)
    step_completed        = Signal(int)         # step_idx
    all_finished          = Signal()
    error                 = Signal(str)

    def __init__(self, engine: MeasurementEngine, settings: dict,
                 start_idx: int = 0,
                 stop_idx: int | None = None,
                 remeasure: bool = False,
                 skip_indices: set[int] | None = None) -> None:
        super().__init__()
        self._engine        = engine
        self._settings      = settings
        self._start_idx     = start_idx
        self._stop_idx      = stop_idx   # if set, only run up to (and including) this step
        self._remeasure     = remeasure  # if True, version new data instead of overwriting
        self._skip_indices  = skip_indices or set()
        self._gate          = threading.Event()
        self._stop          = threading.Event()

    def confirm(self) -> None:
        """모드 전환 완료 — main thread 에서 호출."""
        self._gate.set()

    def stop(self) -> None:
        self._stop.set()
        self._gate.set()          # 대기 중이면 즉시 해제
        self._engine.stop_sequence()

    def run(self) -> None:
        original_cb = self._engine.on_progress
        _SEQ_LABELS = {
            "lum_swing": "Lum. Swing", "lum_loading": "APL Loading",
            "gamut": "Gamut",          "contrast": "Contrast",
        }
        try:
            prev_mode_key: tuple | None = None

            for step_idx, (mode_idx, is_hdr, case, seq_name) in enumerate(_ALL_STEPS):
                if step_idx < self._start_idx:
                    continue
                if step_idx in self._skip_indices:
                    continue
                if self._stop_idx is not None and step_idx > self._stop_idx:
                    break
                if self._stop.is_set():
                    break

                mode     = "HDR" if is_hdr else "SDR"
                key      = f"{mode}_{case}"
                mode_key = (is_hdr, case)

                # 모드가 바뀔 때만 신호 전환 → 확인 다이얼로그 / LG 명령 전송
                if prev_mode_key is None or mode_key != prev_mode_key:
                    # ① 패턴 제너레이터 신호 먼저 전환 (idempotent — 이미 같은 상태면 무시)
                    #    TV 입력단에 HDR 신호가 들어온 뒤 사용자가 PSM 모드를 변경할 수 있다
                    gen = self._engine.generator
                    if gen is not None and gen.is_connected:
                        if is_hdr:
                            gen.set_hdr(True)
                        else:
                            gen.set_sdr()
                    # ② 확인 다이얼로그 / LG 명령 전송
                    self._gate.clear()
                    self.mode_change_requested.emit(is_hdr, case)
                    self._gate.wait()
                    if self._stop.is_set():
                        break
                    prev_mode_key = mode_key

                label = f"{mode} {case} — {_SEQ_LABELS.get(seq_name, seq_name)}"
                self.step_started.emit(step_idx, label)
                self._engine.on_progress = lambda s, p, d, si=step_idx: \
                    self.progress.emit(f"step{si}/{s}", p, d)

                if seq_name == "lum_swing":
                    result = self._engine.run_sequence(
                        "lum_swing", case=case, is_hdr=is_hdr,
                        sample_count=self._settings.get("swing_sample_count", 301),
                        interval_sec=float(self._settings.get("swing_interval_sec", 1.0)),
                    )
                    sk = _versioned_key(self._engine.session_swing, key) if self._remeasure else key
                    self._engine.session_swing[sk] = list(result.get(case, []))

                elif seq_name == "lum_loading":
                    result = self._engine.run_sequence(
                        "lum_loading",
                        version=self._settings.get("version", "37"),
                        case=case, is_hdr=is_hdr,
                        cooling_enabled=self._settings.get("cooling_enabled", False),
                        cooling_apl_threshold=self._settings.get("cooling_apl_threshold", 10),
                        cooling_duration_sec=float(self._settings.get("cooling_duration_sec", 5)),
                        measurements_per_step=self._settings.get("measurements_per_step", 3),
                    )
                    lk = _versioned_key(self._engine.session_loading, key) if self._remeasure else key
                    self._engine.session_loading[lk] = dict(result)

                elif seq_name == "gamut":
                    result = self._engine.run_sequence("gamut", is_hdr=is_hdr)
                    gk = _versioned_key(self._engine.session_gamut, mode) if self._remeasure else mode
                    self._engine.session_gamut[gk] = dict(result)

                elif seq_name == "contrast":
                    result = self._engine.run_sequence("contrast", is_hdr=is_hdr)
                    ck = _versioned_key(self._engine.session_contrast, mode) if self._remeasure else mode
                    self._engine.session_contrast[ck] = dict(result)

                if self._stop.is_set():
                    break

                self.step_completed.emit(step_idx)

            self.all_finished.emit()
        except Exception as exc:
            self.error.emit(str(exc))
        finally:
            self._engine.on_progress = original_cb
