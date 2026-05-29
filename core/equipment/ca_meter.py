from __future__ import annotations

import re
import time
from typing import Optional

import serial

from .base import MeterBase, MeasureResult, PatternInfo

# ---------------------------------------------------------------------------
# CA-410 serial parameters (Konica Minolta spec — Linux/Windows 공통)
#   38400 baud, 7 data bits, Even parity, 2 stop bits, RTS/CTS hardware flow
# ---------------------------------------------------------------------------
_CA410_PARAMS: dict = dict(
    baudrate=38400,
    bytesize=serial.SEVENBITS,
    parity=serial.PARITY_EVEN,
    stopbits=serial.STOPBITS_TWO,
    rtscts=True,
    timeout=5.0,
)

# CA-310 — simpler RS-232 setup
_CA310_PARAMS: dict = dict(
    baudrate=9600,
    bytesize=serial.EIGHTBITS,
    parity=serial.PARITY_NONE,
    stopbits=serial.STOPBITS_ONE,
    rtscts=False,
    timeout=5.0,
)


def _xyz_from_Yxy(Y: float, x: float, y: float) -> tuple[float, float, float]:
    """Convert CIE Yxy to XYZ.  Returns (0,0,0) when y==0."""
    if y == 0:
        return 0.0, Y, 0.0
    X = (x / y) * Y
    Z = ((1.0 - x - y) / y) * Y
    return X, Y, Z


class CaMeter(MeterBase):
    """Driver for Konica Minolta CA-310 / CA-410 colorimeters via pyserial.

    Connect sequence (CA-410):
        open port → COM,1 (REMOTE) → MDR,0 (hold release) → IDQ (identify)

    Measure:
        MES → OK,Y,x,y,u',v',CCT,duv  (standard)
             or  OK00,Pxx v1;v2;v3     (compact)

    Disconnect:
        COM,0 (LOCAL) → close port

    Linux  : port = "/dev/ttyUSB0"
    Windows: port = "COM3"
    """

    SUPPORTED_MODELS = ("CA-310", "CA-410")

    _CMD_TIMEOUT = 3.0   # COM,1 / IDQ 응답 대기
    _MES_TIMEOUT = 8.0   # MES 측정 응답 대기 (저휘도 시 느릴 수 있음)
    _MES_WAIT    = 0.3   # MES 전송 후 최소 대기

    def __init__(self, model: str = "CA-410") -> None:
        if model not in self.SUPPORTED_MODELS:
            raise ValueError(f"Unsupported model {model!r}. Choose from {self.SUPPORTED_MODELS}")
        self.model = model
        self._serial: Optional[serial.Serial] = None
        self._port: Optional[str] = None
        self._remote: bool = False
        self.ident: Optional[str] = None
        self._current_pattern: PatternInfo = PatternInfo(
            type="unknown", apl_pct=0.0, width_pct=0.0, height_pct=0.0, color="unknown"
        )

    # ------------------------------------------------------------------
    # MeterBase interface
    # ------------------------------------------------------------------

    def connect(self, port: str) -> None:
        """포트 열기 → REMOTE 전환(COM,1) → 홀드 해제(MDR,0) → ID 조회(IDQ)."""
        if self.is_connected:
            self.disconnect()

        params = _CA410_PARAMS if self.model == "CA-410" else _CA310_PARAMS
        self._serial = serial.Serial(port, **params)
        self._port = port
        time.sleep(0.3)  # DTR/RTS 안정화

        # CA-310은 REMOTE 핸드셰이크 없이 바로 측정 가능
        if self.model == "CA-310":
            return

        # REMOTE 모드 전환
        resp = self._cmd(b"COM,1\r\n", wait=0.3, timeout=self._CMD_TIMEOUT)
        if "OK" not in resp:
            self._serial.close()
            self._serial = None
            raise RuntimeError(
                f"CA-410 REMOTE 전환 실패 (응답: {resp!r})\n"
                "장비 전원 / 케이블 / baud rate(38400 7E2) 확인"
            )
        self._remote = True

        # 홀드 해제 — 미지원 장비는 무시
        self._cmd(b"MDR,0\r\n", wait=0.1, timeout=self._CMD_TIMEOUT)

        # ID 조회
        id_resp = self._cmd(b"IDQ\r\n", wait=0.05, timeout=self._CMD_TIMEOUT)
        if id_resp.startswith("OK"):
            parts = id_resp.split(",")
            self.ident = ",".join(parts[1:]).strip() if len(parts) > 1 else id_resp
        else:
            self.ident = id_resp or "(ID 조회 실패)"

    def disconnect(self) -> None:
        if self._serial and self._serial.is_open:
            if self._remote:
                try:
                    self._cmd(b"COM,0\r\n", wait=0.1, timeout=self._CMD_TIMEOUT)
                except Exception:
                    pass
            self._serial.close()
        self._serial = None
        self._port = None
        self._remote = False
        self.ident = None

    def measure(self) -> MeasureResult:
        if not self._serial or not self._serial.is_open:
            # USB 글리치 등으로 포트가 닫혔을 때 자동 재연결
            if self._port:
                self.connect(self._port)
            else:
                raise RuntimeError("CaMeter is not connected")

        # CA-410: REMOTE 상태가 풀렸으면 재전환
        if self.model == "CA-410" and not self._remote:
            resp = self._cmd(b"COM,1\r\n", wait=0.3, timeout=self._CMD_TIMEOUT)
            if "OK" not in resp:
                raise RuntimeError(f"CA-410 REMOTE 재전환 실패: {resp!r}")
            self._remote = True
            self._cmd(b"MDR,0\r\n", wait=0.1, timeout=self._CMD_TIMEOUT)

        timestamp_ms = int(time.time() * 1000)

        resp = self._cmd(b"MES\r\n", wait=self._MES_WAIT, timeout=self._MES_TIMEOUT)
        if not resp:
            raise TimeoutError("CA 미터 응답 없음 (측정 범위 초과 또는 케이블 확인)")

        return self._parse_mes(resp, timestamp_ms)

    @property
    def is_connected(self) -> bool:
        return self._serial is not None and self._serial.is_open

    # ------------------------------------------------------------------
    # Pattern context helper
    # ------------------------------------------------------------------

    def set_current_pattern(self, pattern_info: PatternInfo) -> None:
        self._current_pattern = pattern_info

    # ------------------------------------------------------------------
    # Response parsing
    # ------------------------------------------------------------------

    def _parse_mes(self, resp: str, timestamp_ms: int) -> MeasureResult:
        """MES 응답 파싱 — 두 가지 포맷 지원.

        표준:   OK,Y,x,y,u',v',CCT,duv
        Compact: OK00,Pxx v1;v2;v3  (v1=x×N, v2=y×N, v3=Y)
        """
        parts = [p.strip() for p in resp.split(",")]
        status = parts[0]

        if status.startswith("NG"):
            code = parts[1] if len(parts) > 1 else "?"
            raise RuntimeError(f"CA 미터 에러 NG,{code} — 측정 조건 확인")

        if not status.startswith("OK"):
            raise RuntimeError(f"CA 미터 알 수 없는 응답: {resp!r}")

        # ── 표준 포맷: OK,Y,x,y,u',v',CCT,duv ─────────────────────────
        if status == "OK" and len(parts) >= 8:
            try:
                Lv   = float(parts[1])
                x    = float(parts[2])
                y    = float(parts[3])
                up   = float(parts[4])
                vp   = float(parts[5])
                cct  = float(parts[6])
                duv  = float(parts[7])
                X, Y, Z = _xyz_from_Yxy(Lv, x, y)
                return MeasureResult(
                    timestamp_ms=timestamp_ms,
                    Lv=Lv, x=x, y=y,
                    u_prime=up, v_prime=vp,
                    X=X, Y=Y, Z=Z,
                    cct=cct, duv=duv,
                    pattern_info=self._current_pattern,
                )
            except ValueError as e:
                raise RuntimeError(f"CA 미터 응답 파싱 실패: {e} | raw={resp!r}")

        # ── Compact 포맷: OK00,Pxx v1;v2;v3 ───────────────────────────
        data_str = ",".join(parts[1:]).strip()
        m = re.match(r"P\d+\s+([\d.]+)\s*;\s*([\d.]+)\s*;\s*([\d.]+)", data_str)
        if m:
            raw1, raw2 = float(m.group(1)), float(m.group(2))
            # x,y 는 CIE 0~1 범위 — 자릿수별 스케일 자동 감지
            if   raw1 <= 1.0:   cx, cy = raw1,            raw2
            elif raw1 < 10.0:   cx, cy = raw1 / 10.0,     raw2 / 10.0
            elif raw1 < 100.0:  cx, cy = raw1 / 100.0,    raw2 / 100.0
            elif raw1 < 1000.0: cx, cy = raw1 / 1000.0,   raw2 / 1000.0
            else:               cx, cy = raw1 / 10000.0,  raw2 / 10000.0
            Lv = float(m.group(3))
            X, Y, Z = _xyz_from_Yxy(Lv, cx, cy)
            denom = -2 * cx + 12 * cy + 3
            up = (4 * cx / denom) if denom != 0 else 0.0
            vp = (9 * cy / denom) if denom != 0 else 0.0
            return MeasureResult(
                timestamp_ms=timestamp_ms,
                Lv=Lv, x=cx, y=cy,
                u_prime=up, v_prime=vp,
                X=X, Y=Y, Z=Z,
                cct=0.0, duv=0.0,
                pattern_info=self._current_pattern,
            )

        raise RuntimeError(f"CA 미터 알 수 없는 응답 포맷: {resp!r}")

    # ------------------------------------------------------------------
    # Low-level serial I/O
    # ------------------------------------------------------------------

    def _cmd(self, cmd: bytes, wait: float = 0.1, timeout: float = 3.0) -> str:
        """명령 전송 → 응답 한 줄 반환."""
        try:
            self._serial.reset_input_buffer()
            self._serial.write(cmd)
            if wait > 0:
                time.sleep(wait)
            self._serial.timeout = timeout
            resp = self._serial.readline().decode("ascii", errors="replace").strip()
            if not resp:
                extra = self._serial.read_all().decode("ascii", errors="replace").strip()
                resp = extra
            return resp
        except Exception as e:
            raise RuntimeError(f"CA 미터 통신 오류: {e}") from e
