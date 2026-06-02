from __future__ import annotations

import time
from typing import Optional

import serial

from .base import GeneratorBase, PatternConfig

# ---------------------------------------------------------------------------
# VG-876 / VG-879 binary serial protocol
#
# Frame format:
#   STX(0x02) + 0xFD + CMD1 + CMD2 + ASCII(p0) + ',' + ASCII(p1) + ... + ETX(0x03)
#
# Response:
#   Success  → STX + ACK(0x06) + ETX
#   Error    → STX + ESTS(0x11) + error_code_ascii + ETX
#   Data     → STX + TRDT(0x10) + data + ETX
#   NAK      → STX + NAK(0x15) + ETX
#
# Key commands:
#   EXPDN4  [0x24 0x20]  execute program / apply settings
#   SOT4    [0x20 0x24]  set output bit length & color format
#   WINDRAW [0x28 0x61]  window area definition (x1,y1,x2,y2)
#   WINCOL4 [0x28 0x62]  window color (r,g,b,bit_mode)
#   SHDMI4  [0x20 0x36]  HDMI output settings
#   SIF4    [0x20 0x38]  AVI InfoFrame
#   SHDR4   [0x20 0xC5]  HDR10 Dynamic Range metadata
# ---------------------------------------------------------------------------

_STX  = b'\x02'
_ETX  = b'\x03'
_ENQ  = b'\x05'
_ACK  = b'\x06'
_NAK  = b'\x15'
_ESTS = b'\x11'
_TRDT = b'\x10'
_SEP  = b'\x2C'   # comma separator

_SCREEN_W = 3840
_SCREEN_H = 2160

# ─── VG-879 프로그램 번호 ─────────────────────────────────────────────────────
#
# _PROG_4K60 : 3840×2160 60p 타이밍 + 컬러바 프리셋.
#              EXPDN4(2286, 0)으로 실행 → 타이밍 안정화 + 컬러바 출력.
#              측정 패턴 출력 전 SDR 초기화에만 사용.
#              이후 _reset_window_memory() → WINDOW4 → WINCOL4 → EXPDN4(0,0) 으로
#              실제 측정 패턴으로 덮어씌운다.
#
# ※ SPTS4 [20H 2AH] 단독 코드(0/1/2 등)는 R/G/B 출력을 비활성화하는 문제가 있음.
#   단, SPTS4(9999, 0,1,2,6,28,26) 조합은 ABC 버튼 + ㅁ + X + R+G+B 동시 활성화
#   가능함이 실험으로 확인됨 → show_center_align() 에서만 이 조합을 사용.
# ─────────────────────────────────────────────────────────────────────────────
_PROG_4K60 = 2286  # 4K60p 타이밍 프로그램 번호

_VG_PARAMS: dict = dict(
    baudrate=38400,
    bytesize=serial.EIGHTBITS,
    parity=serial.PARITY_NONE,
    stopbits=serial.STOPBITS_ONE,
    rtscts=False,
    timeout=1.0,
)


def _build_frame(cmd1: int, cmd2: int, *params) -> bytes:
    parts = _SEP.join(str(p).encode() for p in params)
    return _STX + b'\xFD' + bytes([cmd1, cmd2]) + parts + _ETX


class VgGenerator(GeneratorBase):
    """Driver for Astro Design VG-876 and VG-879 pattern generators.

    Communication is binary-framed ASCII over RS-232 (9600 baud).
    """

    SUPPORTED_MODELS = ("VG-876", "VG-879")

    def __init__(self, model: str = "VG-876") -> None:
        if model not in self.SUPPORTED_MODELS:
            raise ValueError(f"Unsupported model {model!r}. Choose from {self.SUPPORTED_MODELS}")
        self.model = model
        self._serial: Optional[serial.Serial] = None
        self._is_hdr: bool = False
        self._last_pattern: Optional[PatternConfig] = None
        self._timing_loaded: bool = False  # True after EXPDN4(2286,0) — 패턴 간 재로드 방지

    # ------------------------------------------------------------------
    # GeneratorBase interface
    # ------------------------------------------------------------------

    def connect(self, port: str) -> None:
        self._serial = serial.Serial(port, **_VG_PARAMS)
        self._serial.dtr = True
        self._serial.rts = True
        time.sleep(0.5)
        self._enter_terminal_mode()

    def disconnect(self) -> None:
        if self._serial and self._serial.is_open:
            self._serial.close()
        self._serial = None
        self._timing_loaded = False

    def set_pattern(self, cfg: PatternConfig) -> None:
        if cfg.type == "full_field":
            self.show_full_field(cfg.r, cfg.g, cfg.b, cfg.bit_mode)
        elif cfg.type == "raster_window":
            self.show_white_raster_black_window(cfg.width_pct, cfg.bit_mode)
        elif cfg.type == "window":
            self.show_window_patch(
                cfg.width_pct, cfg.height_pct,
                cfg.r, cfg.g, cfg.b,
                cfg.bg_r, cfg.bg_g, cfg.bg_b,
                cfg.bit_mode,
            )
        elif cfg.type == "crosshair":
            self.show_crosshair(cfg.bit_mode)
        else:
            raise ValueError(f"Unknown pattern type: {cfg.type!r}")
        self._last_pattern = cfg

    def set_hdr(self, enabled: bool) -> None:
        if enabled:
            self._setup_hdr10()
        else:
            self.set_sdr()
        self._reapply_last_pattern()

    def set_sdr(self) -> None:
        """4K 60p SDR 출력으로 전환한다."""
        self._timing_loaded = False
        self._load_init_pattern(sleep=1.0)
        # SHDR4 Off: Dynamic Range InfoFrame을 명시적으로 끔
        self._send(_build_frame(
            0x20, 0xC5,
            0, 0,          # Program NO, On/Off=OFF
            7, 1, 0, 0,    # Type, Version, EOTF=SDR, MetaID
            70800, 29200, 17000, 79700, 13100, 4600,
            31270, 32900,
            4000, 1, 4000, 400,  # Max=4000 cd/m², Min=1 (0.0001 nit ×10000)
            0,
        ))
        time.sleep(0.1)
        self._send(_build_frame(0x24, 0x20, 0, 0))
        self._is_hdr = False

    @property
    def is_connected(self) -> bool:
        return self._serial is not None and self._serial.is_open

    # ------------------------------------------------------------------
    # Internal initialization helper
    # ------------------------------------------------------------------

    def _load_init_pattern(self, sleep: float = 1.5) -> None:
        """4K60p 타이밍 로드 (SDR 초기화).

        EXPDN4(2286,0) 만 전송한다. 타이밍 안정화 후 화면에 컬러바가 잠깐
        보이지만, 곧바로 호출되는 _reset_window_memory → WINDOW4 → WINCOL4 →
        EXPDN4(0,0) 시퀀스가 덮어씌우므로 측정 패턴에 영향 없다.

        ※ 일반 패턴 출력에는 SPTS4를 사용하지 않는다.
          센터 정렬 패턴(show_center_align)에서만 확인된 조합으로 사용.
        """
        self._send(_build_frame(0x24, 0x20, _PROG_4K60, 0))  # EXPDN4(2286,0): 타이밍 로드
        time.sleep(sleep)
        self._timing_loaded = True

    def _prepare_pattern_base(self) -> None:
        """패턴 출력 전 타이밍 초기화 (필요할 때만).

        SDR: _timing_loaded=False 일 때만 EXPDN4(2286,0) 로드.
             set_sdr() / reset() 이 이미 로드했으면 재로드 안 함
             → gamut R→G→B 전환 시 컬러바 깜빡임 방지.
        HDR: _setup_hdr10() 이 SHDMI4/SIF4/SHDR4 로 설정했으므로 항상 스킵.
        """
        if not self._is_hdr and not self._timing_loaded:
            self._load_init_pattern(sleep=1.5)

    def _scale_rgb_for_output(
        self,
        r: int,
        g: int,
        b: int,
        bit_mode: int,
    ) -> tuple[int, int, int, int]:
        if self._is_hdr and bit_mode == 8:
            return r * 4 + r // 64, g * 4 + g // 64, b * 4 + b // 64, 10
        return r, g, b, bit_mode

    def _reapply_last_pattern(self) -> None:
        if self._last_pattern is None:
            # 이전 패턴 없음 → HDR/SDR 신호가 활성됐음을 보여주는 중간 회색 출력
            self.show_full_field(128, 128, 128)
            return
        last_pattern = self._last_pattern
        self._last_pattern = None
        self.set_pattern(last_pattern)

    # ------------------------------------------------------------------
    # Pattern output helpers
    # ------------------------------------------------------------------

    def _reset_window_memory(self) -> None:
        """ALLCLR4 [28H 60H] — 모든 플레인(윈도우 + 베이스 패턴) 클리어.

        WINCLR4 [28H 63H] 는 윈도우 플레인만 지우고 컬러바 베이스 패턴이 남아
        BCOL4(black) 배경이 컬러바로 보이는 문제가 있다.
        ALLCLR4 는 All planes 클리어이므로 컬러바 베이스까지 완전 소거 →
        이후 EXPDN4(0,0) + BCOL4(black) 가 실제 블랙 배경을 표시한다.

        장비 freeze 시: UI [장비 리셋] 버튼으로 복구 후 사용.
        """
        self._send(_build_frame(0x28, 0x60))  # ALLCLR4 [28H 60H]
        time.sleep(0.05)

    def reset(self) -> None:
        """장비 freeze/응답불량 시 초기 상태로 복귀한다.

        ENQ → ACK 터미널 모드 재진입 후 4K60p 컬러바(2286번)를 다시 로드한다.
        내부 상태(HDR 플래그, 마지막 패턴)도 초기화한다.
        """
        self._enter_terminal_mode()
        self._load_init_pattern(sleep=2.0)
        self._is_hdr = False
        self._last_pattern = None

    def show_black(self) -> dict:
        """즉시 블랙 출력 — _prepare_pattern_base() 를 거치지 않는 직통 경로.

        쿨링처럼 '지금 바로 화면을 끄고 싶을 때' 사용.
        ALLCLR4 + EXPDN4(0,0) 만 전송해 타이밍 로드 없이 즉시 블랙.
        """
        self._send(_build_frame(0x28, 0x60))        # ALLCLR4: 모든 플레인 클리어
        time.sleep(0.1)
        return self._send(_build_frame(0x24, 0x20, 0, 0))  # EXPDN4(0,0)

    def show_center_align(self) -> dict:
        """ABC 센터 정렬 패턴: VG-879 ABC 버튼 + ㅁ + X + R+G+B 동시 활성.

        SPTS4(9999, 0,1,2,6,28,26) 단일 호출로 장비 패널의
        ABC 버튼, ㅁ(네모) 체크, X(십자) 체크, R/G/B 출력 버튼을 동시에 활성화한다.
        코드 조합은 test_abc.py 스캔으로 확인된 값이다.
        """
        self._prepare_pattern_base()
        r, g, b, bit_mode = self._scale_rgb_for_output(255, 255, 255, 8)

        self._send(_build_frame(0x20, 0x2C, 9999, 1,  r, g, b, bit_mode))  # SPT4 fg=white
        time.sleep(0.12)
        self._send(_build_frame(0x20, 0x2C, 9999, 18, 0, 0, 0, 8))         # SPT4 bg=black
        time.sleep(0.12)
        self._send(_build_frame(0x20, 0x2A, 9999, 0, 1, 2, 6, 28, 26))    # SPTS4: ABC+ㅁ+X+RGB
        time.sleep(0.12)
        return self._send(_build_frame(0x24, 0x20, 9999, 0))               # EXPDN4

    def show_full_field(self, r: int, g: int, b: int, bit_mode: int = 8) -> dict:
        """전체 화면을 단일 색상으로 출력한다. (색재현율 측정용)

        연속 full_field 호출 (gamut R→G→B→W→BK):
          처음 또는 이전 패턴이 다른 타입 → ALLCLR4 + WINDOW4(전체) + WINCOL4 + EXPDN4
          연속 full_field → WINCOL4 + EXPDN4 만 전송 (깜빡임·지연 없음)

        타이밍 설명:
          ALLCLR4 는 컬러바 베이스 포함 전체 플레인 클리어.
          WINCOL4 만으로 색상이 즉시 바뀌므로 컬러바 깜빡임이 없다.
        """
        self._prepare_pattern_base()
        r, g, b, bit_mode = self._scale_rgb_for_output(r, g, b, bit_mode)

        last = self._last_pattern
        if last is None or last.type != "full_field":
            # 처음이거나 이전이 다른 패턴 타입 → 전체 화면 창 등록
            self._reset_window_memory()                                                 # ALLCLR4
            self._send(_build_frame(0x28, 0x61, 0, 0, _SCREEN_W - 1, _SCREEN_H - 1))  # WINDOW4: 전체
            time.sleep(0.1)

        self._send(_build_frame(0x28, 0x62, r, g, b, bit_mode))                        # WINCOL4: 색상 (즉시 전환)
        time.sleep(0.05)
        return self._send(_build_frame(0x24, 0x20, 0, 0))                              # EXPDN4(0,0)

    def show_crosshair(self, bit_mode: int = 8) -> dict:
        """컬러바 위에 흰색 십자선(X) 오버레이. (정렬용)

        show_center_align과 동일 방식:
          WINCLR4(윈도우만 클리어) → WINDOW4(수평) → WINDOW4(수직) → WINCOL4(white) → EXPDN4(0,0).
        컬러바 베이스는 유지됨.
        """
        self._prepare_pattern_base()
        r, g, b, bit_mode = self._scale_rgb_for_output(255, 255, 255, bit_mode)
        cx, cy = _SCREEN_W // 2, _SCREEN_H // 2

        self._send(_build_frame(0x28, 0x63))                                          # WINCLR4: 윈도우만 클리어
        time.sleep(0.2)
        self._send(_build_frame(0x28, 0x61, 0, cy - 1, _SCREEN_W - 1, cy + 1))      # WINDOW4: 수평선
        time.sleep(0.2)
        self._send(_build_frame(0x28, 0x61, cx - 1, 0, cx + 1, _SCREEN_H - 1))      # WINDOW4: 수직선
        time.sleep(0.2)
        self._send(_build_frame(0x28, 0x62, r, g, b, bit_mode))                      # WINCOL4(white): 전체 창 적용
        time.sleep(0.2)
        return self._send(_build_frame(0x24, 0x20, 0, 0))                             # EXPDN4(0,0)

    def show_white_raster_black_window(
        self,
        side_pct: float,
        bit_mode: int = 8,
    ) -> dict:
        """White Raster + centered Black Window. (명암비 측정용)

        확인된 시퀀스 (test_raster.py):
          ALLCLR4
          → SPT4(fg=white) + SPT4(bg=black) + SPTS4(0,1,2,10) + EXPDN4(9999,0)
          → WINDOW4(center, side_pct) + WINCOL4(black) + EXPDN4(0,0)

        side_pct: 검은 창 한 변 크기 (% of screen).  100%=전체, 50%=절반 등.
        """
        self._prepare_pattern_base()
        r, g, b, bm = self._scale_rgb_for_output(255, 255, 255, bit_mode)

        # ── White Raster ──────────────────────────────────────────────────────
        self._reset_window_memory()                                          # ALLCLR4
        self._send(_build_frame(0x20, 0x2C, 9999,  1, r, g, b, bm))        # SPT4 fg=white
        time.sleep(0.05)
        self._send(_build_frame(0x20, 0x2C, 9999, 18, 0, 0, 0, 8))         # SPT4 bg=black
        time.sleep(0.05)
        self._send(_build_frame(0x20, 0x2A, 9999, 0, 1, 2, 10))            # SPTS4(0,1,2,10)
        time.sleep(0.05)
        self._send(_build_frame(0x24, 0x20, 9999, 0))                       # EXPDN4(9999,0)
        time.sleep(0.15)

        # ── Black Center Window ───────────────────────────────────────────────
        w  = max(1, int(_SCREEN_W * side_pct / 100))
        h  = max(1, int(_SCREEN_H * side_pct / 100))
        x1 = (_SCREEN_W - w) // 2
        y1 = (_SCREEN_H - h) // 2
        x2 = x1 + w - 1
        y2 = y1 + h - 1

        self._send(_build_frame(0x28, 0x61, x1, y1, x2, y2))               # WINDOW4(center)
        time.sleep(0.05)
        self._send(_build_frame(0x28, 0x62, 0, 0, 0, 8))                   # WINCOL4(black)
        time.sleep(0.05)
        return self._send(_build_frame(0x24, 0x20, 0, 0))                   # EXPDN4(0,0)

    def show_window_patch(
        self,
        width_pct: float,
        height_pct: float,
        r: int, g: int, b: int,
        bg_r: int = 0, bg_g: int = 0, bg_b: int = 0,
        bit_mode: int = 8,
    ) -> dict:
        """중앙 정렬 윈도우 패치를 출력한다. (휘도 스윙/로딩/명암비 측정용)

        커맨드 순서 — 배경색에 따라 두 방식:

        [어두운 배경 (bg=black)] — lum_loading / lum_swing / contrast 100%
          ALLCLR4 → WINDOW4(중앙) → WINCOL4(밝은색) → EXPDN4(0,0)
          ALLCLR4 기본값이 블랙이므로 중앙 이외 영역은 자동으로 블랙.

        [밝은 배경 (bg≠black)] — contrast 50/20/14.1%
          ALLCLR4 → WINDOW4(상/하/좌/우 4구역) → WINCOL4(밝은 배경색) → EXPDN4(0,0)
          BCOL4는 ALLCLR4 이후 효과 없음. 배경을 4 스트립 창으로 등록하고
          중앙은 미등록(=ALLCLR4 기본 블랙)으로 둔다.
        """
        self._prepare_pattern_base()
        out_bit_mode = 10 if self._is_hdr and bit_mode == 8 else bit_mode
        r, g, b, _ = self._scale_rgb_for_output(r, g, b, bit_mode)
        bg_r, bg_g, bg_b, _ = self._scale_rgb_for_output(bg_r, bg_g, bg_b, bit_mode)
        bit_mode = out_bit_mode

        w = max(1, int(_SCREEN_W * width_pct / 100))
        h = max(1, int(_SCREEN_H * height_pct / 100))
        x1 = (_SCREEN_W - w) // 2
        y1 = (_SCREEN_H - h) // 2
        x2 = x1 + w - 1
        y2 = y1 + h - 1

        self._reset_window_memory()  # ALLCLR4

        if bg_r == 0 and bg_g == 0 and bg_b == 0:
            # 어두운 배경: 중앙 창만 등록, 배경은 ALLCLR4 기본 블랙
            self._send(_build_frame(0x28, 0x61, x1, y1, x2, y2))            # WINDOW4: 중앙
            time.sleep(0.03)
            self._send(_build_frame(0x28, 0x62, r, g, b, bit_mode))         # WINCOL4: 창 색
            time.sleep(0.03)
        else:
            # 밝은 배경: 4 스트립으로 배경 등록, 중앙은 미등록(블랙)
            # WINCOL4는 전역이므로 마지막 1회로 모든 스트립에 적용됨
            if y1 > 0:
                self._send(_build_frame(0x28, 0x61, 0, 0, _SCREEN_W - 1, y1 - 1))          # 상단
                time.sleep(0.1)
            if y2 < _SCREEN_H - 1:
                self._send(_build_frame(0x28, 0x61, 0, y2 + 1, _SCREEN_W - 1, _SCREEN_H - 1))  # 하단
                time.sleep(0.1)
            if x1 > 0:
                self._send(_build_frame(0x28, 0x61, 0, y1, x1 - 1, y2))                    # 좌측
                time.sleep(0.1)
            if x2 < _SCREEN_W - 1:
                self._send(_build_frame(0x28, 0x61, x2 + 1, y1, _SCREEN_W - 1, y2))        # 우측
                time.sleep(0.1)
            self._send(_build_frame(0x28, 0x62, bg_r, bg_g, bg_b, bit_mode))               # WINCOL4: 배경색
            time.sleep(0.2)

        return self._send(_build_frame(0x24, 0x20, 0, 0))                   # EXPDN4(0,0): 렌더링
    # ------------------------------------------------------------------
    # HDR10 setup sequence (vg_hdr.py 기반)
    # ------------------------------------------------------------------

    def _setup_hdr10(
        self,
        max_lum: int = 4000,
        min_lum_nit: float = 0.0001,
        max_cll: int = 4000,
        max_fall: int = 400,
        hdr_on: bool = True,
    ) -> None:
        # Step 1: 4K60p 타이밍 로드 — SHDR4보다 반드시 먼저 (이후 EXPDN4(2286,0) 시 HDR 설정 리셋됨)
        self._send(_build_frame(0x24, 0x20, _PROG_4K60, 0))
        time.sleep(1.0)

        # Step 2: SHDMI4 [20H 36H] — YCbCr422, 10-bit, BT.2020
        self._send(_build_frame(0x20, 0x36, 0, 1, 1, 0, 2, 0, 0))
        time.sleep(0.3)

        # Step 3: SHDR4 [20H C5H] — Dynamic Range and Mastering InfoFrame
        # Max Disp Mastering : 단위 1 cd/m²         (0-65535) → 4000 nits = 4000
        # Min Disp Mastering : 단위 0.0001 cd/m²    (0-65535, ×10000) → 0.0001 nit = 1
        # Disp Primaries     : 실제값 × 100,000 (BT.2020)
        max_m = max_lum
        min_m = max(1, int(min_lum_nit * 10000))
        on_off = 1 if hdr_on else 0

        self._send(_build_frame(
            0x20, 0xC5,
            0,            # Program NO
            on_off,       # On/Off
            7,            # Type
            1,            # Version
            2,            # EOTF = SMPTE ST2084
            0,            # Metadata ID
            70800, 29200, 17000, 79700, 13100, 4600,  # BT.2020 Disp Primaries ×100000
            31270, 32900,                              # D65 White Point ×100000
            max_m, min_m, max_cll, max_fall,
            0,            # Data Type = HDMI
        ))
        time.sleep(0.1)

        # Step 4: 설정 반영
        self._send(_build_frame(0x24, 0x20, 0, 0))
        self._is_hdr = hdr_on

    # ------------------------------------------------------------------
    # Internal helpers
    # ------------------------------------------------------------------

    def _enter_terminal_mode(self, retries: int = 3, wait: float = 0.3) -> bool:
        """ENQ → ACK handshake to enter Terminal Mode."""
        if not self.is_connected:
            return False
        for _ in range(retries):
            try:
                self._serial.reset_input_buffer()
                self._serial.write(_ENQ)
                time.sleep(wait)
                if self._serial.in_waiting > 0:
                    res = self._serial.read(1)
                    if res == _ACK:
                        return True
            except Exception:
                return False
        return False

    def _send(self, frame: bytes) -> dict:
        if not self.is_connected:
            raise RuntimeError("VgGenerator is not connected")
        self._serial.reset_input_buffer()
        self._serial.write(frame)
        return self._read_response()

    def _read_response(self) -> dict:
        buf = bytearray()
        deadline = time.time() + self._serial.timeout
        try:
            while time.time() < deadline:
                if self._serial.in_waiting > 0:
                    ch = self._serial.read(1)
                    buf.extend(ch)

                    if ch == _ACK:
                        tail = self._serial.read(1)
                        if tail:
                            buf.extend(tail)
                        return {'status': 'ok', 'raw': bytes(buf)}

                    elif ch == _ESTS:
                        err_buf = bytearray()
                        t = time.time() + 0.3
                        while time.time() < t:
                            if self._serial.in_waiting > 0:
                                c = self._serial.read(1)
                                buf.extend(c)
                                if c == _ETX:
                                    break
                                err_buf.extend(c)
                        return {'status': 'ests', 'ests_code': err_buf.decode('ascii', errors='replace'), 'raw': bytes(buf)}

                    elif ch == _TRDT:
                        data_buf = bytearray()
                        t = time.time() + 0.5
                        while time.time() < t:
                            if self._serial.in_waiting > 0:
                                c = self._serial.read(1)
                                buf.extend(c)
                                if c == _ETX:
                                    break
                                data_buf.extend(c)
                        return {'status': 'data', 'data': bytes(data_buf), 'raw': bytes(buf)}

                    elif ch == _NAK:
                        tail = self._serial.read(1)
                        if tail:
                            buf.extend(tail)
                        return {'status': 'nak', 'raw': bytes(buf)}

                    elif ch == _ETX:
                        return {'status': 'data', 'raw': bytes(buf)}

                time.sleep(0.01)
            return {'status': 'timeout', 'raw': bytes(buf)}
        except Exception as e:
            return {'status': 'error', 'error': str(e)}
