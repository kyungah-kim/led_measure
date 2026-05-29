from __future__ import annotations

from typing import Any, Callable, Optional

from PySide6.QtCore import QThread, Signal

from core.engine import MeasurementEngine


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
