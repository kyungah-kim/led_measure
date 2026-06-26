from __future__ import annotations

import threading
from typing import Any, Callable, Dict, List, Optional

from .equipment.base import GeneratorBase, MeterBase, MeasureResult
from .equipment.ca_meter import CaMeter
from .equipment.vg_generator import VgGenerator
from .sequences.center_align import CenterAlignSequence
from .sequences.lum_swing import LumSwingSequence
from .sequences.lum_loading import LumLoadingSequence
from .sequences.gamut import GamutSequence
from .sequences.contrast import ContrastSequence
from .sequences.module_measure import ModuleMeasureSequence, DEFAULT_GAMMA_STEPS
from .sequences.calman_sweep import CalmanSweepSequence, GAMUT_NAMES


class MeasurementEngine:
    """Central orchestrator for all measurement sequences.

    Holds references to the connected meter and generator and dispatches
    sequence runs via run_sequence().  A threading.Lock serialises meter
    access so multiple callers do not collide mid-measurement.

    Design:
    - Instantiate once per application session (singleton-friendly).
    - Pass on_progress callback at construction or per run_sequence call.
    - UI layers (desktop/web) import and call this class; this module must
      never import from desktop/ or web/.
    """

    def __init__(
        self,
        brand: str = "",
        model_name: str = "",
        on_progress: Optional[Callable[[str, float, Any], None]] = None,
    ) -> None:
        self.brand = brand
        self.model_name = model_name
        self.on_progress: Callable[[str, float, Any], None] = on_progress or (
            lambda step, pct, data: None
        )

        self.meter: Optional[MeterBase] = None
        self.generator: Optional[GeneratorBase] = None
        self.lg_tv_serial: Any = None   # LG TV 시리얼 (ConnectionPanel에서 열고, AutoAllPanel에서 사용)
        self.lg_log_tx: Any = None      # LG 터미널 [TX] 로그 콜백 (CenterAlignPanel에서 설정)
        self.auto_save_dir: str = ""  # 공통 자동 저장 폴더 (연결 패널에서 설정)

        # ── LG TV 장치 정보 (연결 시 luna 명령으로 자동 수집) ─────────────────────
        self.lg_serial_number: str = ""
        self.lg_sw_version: str = ""    # core_os_release
        self.lg_sw_codename: str = ""   # core_os_release_codename
        # UI 콜백: (brand, model) → ConnectionPanel이 설정, CenterAlignPanel에서 호출
        self.on_lg_device_info: Optional[Callable[[str, str], None]] = None

        # ── 세션 데이터 (통합 파일 저장용) ──────────────────────────────────────
        # 앱 실행 중 모든 측정 결과를 누적 보관.  패널 _on_finished 때마다 갱신.
        self.session_swing:    Dict[str, Any] = {}   # "SDR_Vivid" → [MeasureResult]
        self.session_loading:  Dict[str, Any] = {}   # "SDR_Vivid" → {apl → results}
        self.session_gamut:    Dict[str, Any] = {}   # "SDR" / "HDR" → {color → result}
        self.session_contrast: Dict[str, Any] = {}   # "SDR" / "HDR" → {side → result}
        self.session_key: str = ""  # "{brand}_{model}" of the model that owns session data

        # Shared lock for meter access across threads
        self.meter_lock = threading.Lock()

        # Currently running sequence (supports stop_sequence())
        self._current_seq: Any = None

    # ------------------------------------------------------------------
    # Connection management
    # ------------------------------------------------------------------

    def connect_meter(self, port: str, model: str = "CA-410") -> None:
        """Instantiate and connect a CA meter."""
        m = CaMeter(model=model)
        m.connect(port)
        self.meter = m

    def connect_generator(self, port: str, model: str = "VG-876") -> None:
        """Instantiate and connect a VG pattern generator."""
        g = VgGenerator(model=model)
        g.connect(port)
        self.generator = g

    def stop_sequence(self) -> None:
        """Signal the currently running sequence to stop."""
        if self._current_seq is not None and hasattr(self._current_seq, "stop"):
            self._current_seq.stop()

    def disconnect_all(self) -> None:
        if self.meter and self.meter.is_connected:
            self.meter.disconnect()
        if self.generator and self.generator.is_connected:
            self.generator.disconnect()

    @property
    def is_ready(self) -> bool:
        """True when both meter and generator are connected."""
        return (
            self.meter is not None
            and self.meter.is_connected
            and self.generator is not None
            and self.generator.is_connected
        )

    # ------------------------------------------------------------------
    # Sequence dispatcher
    # ------------------------------------------------------------------

    def run_sequence(self, seq_name: str, **kwargs: Any) -> Any:
        """Dispatch to a named sequence and return its results.

        seq_name values: "center_align", "lum_swing", "lum_loading",
                         "gamut", "contrast"

        kwargs are forwarded directly to the sequence's run() method.
        A wrapped callback is injected to fire on_progress automatically.
        """
        generator_only = {"center_align"}
        if seq_name in generator_only:
            if self.generator is None or not self.generator.is_connected:
                raise RuntimeError("Generator is not connected")
        elif not self.is_ready:
            raise RuntimeError("Engine is not ready: meter or generator not connected")

        dispatch: Dict[str, Callable[..., Any]] = {
            "center_align":   self._run_center_align,
            "lum_swing":      self._run_lum_swing,
            "lum_loading":    self._run_lum_loading,
            "gamut":          self._run_gamut,
            "contrast":       self._run_contrast,
            "module_measure": self._run_module_measure,
            "calman_sweep":   self._run_calman_sweep,
        }
        runner = dispatch.get(seq_name)
        if runner is None:
            raise ValueError(f"Unknown sequence {seq_name!r}")
        return runner(**kwargs)

    # ------------------------------------------------------------------
    # Private sequence runners
    # ------------------------------------------------------------------

    def _run_center_align(self, **kwargs: Any) -> CenterAlignSequence:
        seq = CenterAlignSequence(self)
        on_ready = kwargs.get("on_ready")
        seq.start(on_ready=on_ready)
        self.on_progress("center_align", 0.0, seq)
        return seq

    def _run_lum_swing(self, **kwargs: Any) -> Dict[str, List[MeasureResult]]:
        case: str = kwargs.get("case", "A")
        is_hdr: bool = kwargs.get("is_hdr", False)
        sample_count: int = kwargs.get("sample_count", LumSwingSequence.DEFAULT_SAMPLE_COUNT)
        interval_sec: float = kwargs.get("interval_sec", LumSwingSequence.DEFAULT_INTERVAL_SEC)
        user_callback: Callable = kwargs.get("callback", lambda r: None)

        seq = LumSwingSequence(self)
        self._current_seq = seq
        collected: List[MeasureResult] = []

        def _cb(result: MeasureResult) -> None:
            collected.append(result)
            progress = len(collected) / sample_count
            self.on_progress("lum_swing", progress, result)
            user_callback(result)

        try:
            results = seq.run(case=case, is_hdr=is_hdr, callback=_cb,
                              sample_count=sample_count, interval_sec=interval_sec)
        finally:
            self._current_seq = None
        self.on_progress("lum_swing", 1.0, results)
        return {case: results}

    def _run_lum_loading(self, **kwargs: Any) -> Dict[str, Any]:
        version: str = str(kwargs.get("version", "37"))
        case: str = kwargs.get("case", "A")
        is_hdr: bool = kwargs.get("is_hdr", False)
        cooling_enabled: bool = kwargs.get("cooling_enabled", False)
        cooling_apl_threshold: int = kwargs.get("cooling_apl_threshold", 10)
        cooling_duration_sec: float = kwargs.get("cooling_duration_sec", 5.0)
        measurements_per_step: int = kwargs.get("measurements_per_step", 5)
        user_callback: Callable = kwargs.get("callback", lambda i, a, r: None)

        from .sequences.lum_loading import _STEP_VERSIONS
        total_steps = len(_STEP_VERSIONS.get(version, []))
        completed = [0]

        seq = LumLoadingSequence(self)
        self._current_seq = seq

        def _cb(step_idx: int, apl: float, step_results: List[MeasureResult]) -> None:
            completed[0] += 1
            progress = completed[0] / max(total_steps, 1)
            self.on_progress("lum_loading", progress, {"apl": apl, "results": step_results})
            user_callback(step_idx, apl, step_results)

        results = seq.run(
            version=version,
            case=case,
            is_hdr=is_hdr,
            cooling_enabled=cooling_enabled,
            callback=_cb,
            measurements_per_step=measurements_per_step,
            cooling_duration_sec=cooling_duration_sec,
            cooling_apl_threshold=cooling_apl_threshold,
        )
        self.on_progress("lum_loading", 1.0, results)
        return results

    def _run_gamut(self, **kwargs: Any) -> Dict[str, MeasureResult]:
        is_hdr: bool = kwargs.get("is_hdr", False)
        user_callback: Callable = kwargs.get("callback", lambda c, r: None)

        # 6 steps total: 1 init + 5 colours
        _TOTAL = 6
        completed = [0]
        seq = GamutSequence(self)
        self._current_seq = seq

        def _cb(color: str, result: MeasureResult) -> None:
            completed[0] += 1
            self.on_progress("gamut", completed[0] / _TOTAL, {"color": color, "result": result})
            user_callback(color, result)

        # 초기화 단계를 progress에 반영
        if is_hdr:
            self.generator.set_hdr(True)
        else:
            self.generator.set_sdr()
        self.on_progress("gamut", 1 / _TOTAL, {"color": None, "result": None})

        results = seq.run(is_hdr=is_hdr, callback=_cb, skip_init=True)
        self.on_progress("gamut", 1.0, results)
        return results

    def _run_module_measure(self, **kwargs: Any) -> Dict[str, Any]:
        is_hdr: bool = kwargs.get("is_hdr", False)
        gamma_channels: List[str] = kwargs.get("gamma_channels", ["W", "R", "G", "B"])
        gamma_steps: List[int] = kwargs.get("gamma_steps", list(DEFAULT_GAMMA_STEPS))
        ref_uv = kwargs.get("ref_uv", {})
        run_gamma: bool = kwargs.get("run_gamma", True)
        run_colors: bool = kwargs.get("run_colors", True)

        gamma_count = len(gamma_channels) * len(gamma_steps) if run_gamma else 0
        color_count = 7 if run_colors else 0
        total = gamma_count + color_count
        completed = [0]

        seq = ModuleMeasureSequence(self)
        self._current_seq = seq

        def _cb(step_name: str, data: Any) -> None:
            if step_name in ("gamma", "color"):
                completed[0] += 1
            pct = min(completed[0] / max(total, 1), 1.0)
            self.on_progress(f"module_{step_name}", pct, data)

        try:
            result = seq.run(
                is_hdr=is_hdr,
                gamma_channels=gamma_channels,
                gamma_steps=gamma_steps,
                ref_uv=ref_uv,
                callback=_cb,
                run_gamma=run_gamma,
                run_colors=run_colors,
            )
        finally:
            self._current_seq = None
        self.on_progress("module_measure", 1.0, result)
        return result

    def _run_contrast(self, **kwargs: Any) -> Dict[float, MeasureResult]:
        is_hdr: bool = kwargs.get("is_hdr", False)
        user_callback: Callable = kwargs.get("callback", lambda s, r: None)

        from .sequences.contrast import _WIN_SIDES_PCT
        total = 1 + len(_WIN_SIDES_PCT)  # full white + black window steps
        completed = [0]
        seq = ContrastSequence(self)
        self._current_seq = seq

        def _cb(win_size: float, result: MeasureResult) -> None:
            completed[0] += 1
            self.on_progress("contrast", completed[0] / total, {"win_size": win_size, "result": result})
            user_callback(win_size, result)

        results = seq.run(is_hdr=is_hdr, callback=_cb)
        self.on_progress("contrast", 1.0, results)
        return results

    def _run_calman_sweep(self, **kwargs: Any) -> Dict[str, Any]:
        from .sequences.calman_sweep import CalmanSweepSequence, COLOR_ORDER, SAT_LEVELS
        is_hdr: bool = kwargs.get("is_hdr", False)
        measured_colors: Dict[str, Any] = kwargs.get("measured_colors", {})

        total = len(COLOR_ORDER) * len(SAT_LEVELS)
        completed = [0]
        seq = CalmanSweepSequence(self)
        self._current_seq = seq

        def _cb(color: str, sat: int, result: MeasureResult, de: float) -> None:
            completed[0] += 1
            pct = completed[0] / total
            self.on_progress("calman_sweep", pct,
                             {"color": color, "sat": sat, "result": result, "de76": de})

        try:
            results = seq.run(
                is_hdr=is_hdr,
                measured_colors=measured_colors,
                callback=_cb,
            )
        finally:
            self._current_seq = None
        self.on_progress("calman_sweep", 1.0, results)
        return results
