from __future__ import annotations

import re
import time
from typing import Optional, Any

import serial

try:
    import usb.core
    import usb.util
    _HAS_PYUSB = True
except ImportError:
    _HAS_PYUSB = False

from .base import MeterBase, MeasureResult, PatternInfo
from ..colorimetry import xy_to_cct_duv

# ---------------------------------------------------------------------------
# CA-310 USB — Konica Minolta COLOR ANALYZER (VID:PID 0686:1002)
#   Interface class 0x00 (vendor-specific), bulk transfer endpoints
# ---------------------------------------------------------------------------
_CA310_USB_VID = 0x0686
_CA310_USB_PID = 0x1002
_CA310_EP_OUT  = 0x02   # Bulk OUT — 명령 전송 (EP2, 64-byte)
_CA310_EP_IN   = 0x82   # Bulk IN  — 응답 수신 (EP2, 64-byte, 공백 패딩)

# ---------------------------------------------------------------------------
# CA-410 serial parameters (Konica Minolta spec)
#   7 data bits, Even parity, 2 stop bits, RTS/CTS hardware flow
# ---------------------------------------------------------------------------
_CA410_PARAMS: dict = dict(
    baudrate=15200,
    bytesize=serial.SEVENBITS,
    parity=serial.PARITY_EVEN,
    stopbits=serial.STOPBITS_TWO,
    rtscts=True,
    timeout=5.0,
)

# CA-310 RS-232C — 38400 bps, 7E2, RTS/CTS 없음
# (USB 연결 시에는 사용하지 않음 — USB 직접 통신 사용)
_CA310_PARAMS: dict = dict(
    baudrate=38400,
    bytesize=serial.SEVENBITS,
    parity=serial.PARITY_EVEN,
    stopbits=serial.STOPBITS_TWO,
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
    """Driver for Konica Minolta CA-310 / CA-410 colorimeters.

    CA-310 연결 방식:
        USB  : port="usb"  → pyusb bulk transfer (VID=0686, PID=1002)
        RS232: port="/dev/ttyUSB1" 등 → pyserial 38400 7E2

    CA-410 연결 방식:
        RS232: port="/dev/ttyUSB0" 등 → pyserial 15200 7E2 + RTS/CTS

    Measure:
        MES → OK,Y,x,y,u',v',CCT,duv  (standard)
    """

    SUPPORTED_MODELS = ("CA-310", "CA-410")

    _CMD_TIMEOUT    = 3.0   # COM,1 / IDQ 응답 대기
    _MES_TIMEOUT    = 10.0  # MES 측정 응답 대기
    _MES_WAIT       = 0.1   # CA-410 MES 전송 후 대기
    _MES_WAIT_CA310 = 0.3   # CA-310 MES 전송 후 대기

    def __init__(self, model: str = "CA-410") -> None:
        if model not in self.SUPPORTED_MODELS:
            raise ValueError(f"Unsupported model {model!r}. Choose from {self.SUPPORTED_MODELS}")
        self.model    = model
        self._serial: Optional[serial.Serial] = None
        self._usb_dev: Optional[Any] = None   # usb.core.Device
        self._port:   Optional[str]  = None
        self._remote: bool           = False
        self.ident:   Optional[str]  = None
        self._current_pattern: PatternInfo = PatternInfo(
            type="unknown", apl_pct=0.0, width_pct=0.0, height_pct=0.0, color="unknown"
        )

    # ------------------------------------------------------------------
    # MeterBase interface
    # ------------------------------------------------------------------

    def connect(self, port: str) -> None:
        """포트 열기 → 모델별 초기화.

        port="usb"  → CA-310 USB 직접 연결 (pyusb)
        port="/dev/ttyUSB..." 등 → 시리얼 연결 (pyserial)
        """
        if self.is_connected:
            self.disconnect()

        self._port = port

        if self.model == "CA-310" and port.lower() == "usb":
            self._connect_ca310_usb()
        elif self.model == "CA-410":
            self._serial = serial.Serial(port, **_CA410_PARAMS)
            time.sleep(0.5)
            self._connect_ca410()
        else:
            # CA-310 serial fallback
            self._serial = serial.Serial(port, **_CA310_PARAMS)
            time.sleep(0.5)
            self._connect_ca310_serial()

    # ------------------------------------------------------------------
    # CA-310 USB (pyusb)
    # ------------------------------------------------------------------

    def _connect_ca310_usb(self) -> None:
        """CA-310 USB 직접 연결 — pyusb bulk transfer."""
        if not _HAS_PYUSB:
            raise RuntimeError("pyusb 가 설치되지 않았습니다: pip install pyusb")

        dev: Any = usb.core.find(idVendor=_CA310_USB_VID, idProduct=_CA310_USB_PID)  # type: ignore[assignment]
        if dev is None:
            raise RuntimeError(
                f"CA-310 USB 장치를 찾을 수 없습니다 (VID={_CA310_USB_VID:04X}:{_CA310_USB_PID:04X})\n"
                "USB 케이블 연결 및 장치 전원을 확인하세요."
            )

        dev = self._open_ca310_dev(dev)

        self._usb_dev = dev

        # REMOTE 모드 전환 — 응답: 'OK00'
        resp = self._cmd_usb(b"COM,1\r\n", timeout=self._CMD_TIMEOUT)
        if "OK" not in resp:
            self._usb_dev = None
            raise RuntimeError(f"CA-310 USB REMOTE 전환 실패 (응답: {resp!r})")

        self._remote = True
        self.ident = "CA-310"

        # MDS,0 — 측정 모드를 CIE1931 xyLv 로 고정
        self._cmd_usb(b"MDS,0\r\n", timeout=self._CMD_TIMEOUT)

    def _open_ca310_dev(self, dev: Any) -> Any:
        """set_configuration() 실행 후 dev 반환.

        리셋 없이 먼저 시도 → COM,1 이 OK00 을 반환하면 성공.
        응답 없으면 USB reset → 재탐색 → 재시도 (cold-start 방어).
        reset 은 마지막 수단으로만 사용해 같은 USB 허브의 다른 장치(패턴 제너레이터)에
        주는 영향을 최소화한다.
        """
        def _setup(d: Any) -> None:
            try:
                if d.is_kernel_driver_active(0):
                    d.detach_kernel_driver(0)
            except Exception:
                pass
            try:
                d.set_configuration()
            except Exception as e:
                if "13" in str(e) or "Access" in str(e) or "Permission" in str(e):
                    import platform
                    if platform.system() == "Windows":
                        guide = (
                            "Windows 설정: Zadig(https://zadig.akeo.ie) 실행 →\n"
                            "Options → List All Devices → COLOR ANALYZER 선택 →\n"
                            "Driver: WinUSB → Replace Driver"
                        )
                    else:
                        guide = (
                            "Linux 설정:\n"
                            "sudo sh -c 'echo SUBSYSTEM==\"usb\", ATTR{idVendor}==\"0686\", "
                            "ATTR{idProduct}==\"1002\", MODE=\"0666\", GROUP=\"plugdev\" "
                            "> /etc/udev/rules.d/99-ca310.rules'\n"
                            "sudo udevadm control --reload-rules && sudo udevadm trigger\n"
                            "이후 CA-310 USB를 뽑았다가 다시 연결하세요."
                        )
                    raise RuntimeError(f"CA-310 USB 접근 권한 없음\n\n{guide}") from e
                raise

        # 1차 시도 — 리셋 없이
        _setup(dev)
        time.sleep(0.5)
        # 64-byte 패딩 COM,1 으로 응답 확인
        padded = b"COM,1\r\n" + b"\x00" * (64 - len(b"COM,1\r\n"))
        try:
            dev.write(_CA310_EP_OUT, padded, timeout=2000)
            time.sleep(0.2)
            raw = bytes(dev.read(_CA310_EP_IN, 512, timeout=2000))
            if b"OK" in raw:
                return dev  # 리셋 불필요, 바로 사용 가능
        except Exception:
            pass

        # 2차 시도 — USB reset 후 재탐색 (cold-start 대응)
        try:
            dev.reset()
        except Exception:
            pass
        time.sleep(1.5)

        dev = usb.core.find(idVendor=_CA310_USB_VID, idProduct=_CA310_USB_PID)  # type: ignore[assignment]
        if dev is None:
            raise RuntimeError("CA-310 USB reset 후 장치 재탐색 실패")

        _setup(dev)
        time.sleep(0.5)
        return dev

    def _cmd_usb(self, cmd: bytes, timeout: float = 3.0) -> str:
        """USB bulk transfer — EP2(0x02) 전송, EP2(0x82) 수신.

        CA-310은 64바이트 패킷을 요구함 — 명령을 null로 패딩하여 전송.
        응답은 공백 패딩된 고정 블록.
        """
        assert self._usb_dev is not None
        # 64바이트 패딩 (CA-310 필수 요건)
        padded = cmd + b"\x00" * max(0, 64 - len(cmd))
        try:
            self._usb_dev.write(_CA310_EP_OUT, padded, timeout=int(timeout * 1000))
            time.sleep(0.2)
            raw = bytes(self._usb_dev.read(_CA310_EP_IN, 512, timeout=int(timeout * 1000)))
            return raw.decode("ascii", errors="replace").strip()
        except usb.core.USBTimeoutError:
            return ""
        except Exception as e:
            raise RuntimeError(f"CA-310 USB 통신 오류: {e}") from e

    # ------------------------------------------------------------------
    # CA-310 serial (RS-232C fallback)
    # ------------------------------------------------------------------

    def _connect_ca310_serial(self) -> None:
        """CA-310 시리얼 연결 — 38400 7E2."""
        assert self._serial is not None
        resp = self._cmd(b"COM,1\r\n", wait=0.3, timeout=self._CMD_TIMEOUT)
        if "OK" not in resp:
            self._serial.close()
            self._serial = None
            raise RuntimeError(
                f"CA-310 REMOTE 전환 실패 (응답: {resp!r})\n"
                "38400 7E2, RTS/CTS 없음 — 케이블/전원/포트 확인"
            )
        self._remote = True

        try:
            id_resp = self._cmd(b"IDQ\r\n", wait=0.3, timeout=self._CMD_TIMEOUT)
            if id_resp.startswith("OK"):
                parts = id_resp.split(",")
                self.ident = ",".join(parts[1:]).strip() if len(parts) > 1 else id_resp
            else:
                self.ident = id_resp or "CA-310"
        except Exception:
            self.ident = "CA-310"

    # ------------------------------------------------------------------
    # CA-410
    # ------------------------------------------------------------------

    def _connect_ca410(self) -> None:
        """CA-410 전용 — COM,1 REMOTE + MDR,0 홀드 해제 + IDQ."""
        resp = self._cmd(b"COM,1\r\n", wait=0.3, timeout=self._CMD_TIMEOUT)
        if "OK" not in resp:
            if self._serial:
                self._serial.close()
            self._serial = None
            raise RuntimeError(
                f"CA-410 REMOTE 전환 실패 (응답: {resp!r})\n"
                "장비 전원 / 케이블 / baud rate(15200 7E2 + RTS/CTS) 확인"
            )
        self._remote = True
        self._cmd(b"MDR,0\r\n", wait=0.1, timeout=self._CMD_TIMEOUT)

        id_resp = self._cmd(b"IDQ\r\n", wait=0.05, timeout=self._CMD_TIMEOUT)
        if id_resp.startswith("OK"):
            parts = id_resp.split(",")
            self.ident = ",".join(parts[1:]).strip() if len(parts) > 1 else id_resp
        else:
            self.ident = id_resp or "(ID 조회 실패)"

    # ------------------------------------------------------------------
    # disconnect / measure / is_connected
    # ------------------------------------------------------------------

    def disconnect(self) -> None:
        try:
            if self._usb_dev is not None:
                if self._remote:
                    try:
                        self._cmd_usb(b"COM,0\r\n", timeout=1.0)
                    except Exception:
                        pass
                usb.util.dispose_resources(self._usb_dev)
            elif self._serial and self._serial.is_open:
                if self._remote:
                    try:
                        self._cmd(b"COM,0\r\n", wait=0.1, timeout=1.0)
                    except Exception:
                        pass
                self._serial.close()
        except Exception:
            pass
        self._usb_dev = None
        self._serial  = None
        self._port    = None
        self._remote  = False
        self.ident    = None

    def measure(self) -> MeasureResult:
        if not self.is_connected:
            if self._port:
                self.connect(self._port)
            else:
                raise RuntimeError("CaMeter is not connected")

        # REMOTE 상태가 풀렸으면 재전환
        if not self._remote:
            resp = self._send_cmd(b"COM,1\r\n")
            if "OK" not in resp:
                raise RuntimeError(f"{self.model} REMOTE 재전환 실패: {resp!r}")
            self._remote = True
            if self.model == "CA-410":
                try:
                    self._send_cmd(b"MDR,0\r\n")
                except Exception:
                    pass

        timestamp_ms = int(time.time() * 1000)

        mes_wait = self._MES_WAIT_CA310 if self.model == "CA-310" else self._MES_WAIT
        resp = ""
        for _attempt in range(3):
            if self._usb_dev is not None:
                resp = self._cmd_usb(b"MES\r\n", timeout=self._MES_TIMEOUT)
            else:
                resp = self._cmd(b"MES\r\n", wait=mes_wait, timeout=self._MES_TIMEOUT)
            if resp:
                break
            if self._port:
                try:
                    self.connect(self._port)
                except Exception:
                    pass
            time.sleep(0.5)

        if not resp:
            raise TimeoutError("CA 미터 응답 없음 (측정 범위 초과 또는 케이블 확인)")

        return self._parse_mes(resp, timestamp_ms)

    @property
    def is_connected(self) -> bool:
        if self._usb_dev is not None:
            return True
        return self._serial is not None and self._serial.is_open

    # ------------------------------------------------------------------
    # Unified send helper
    # ------------------------------------------------------------------

    def _send_cmd(self, cmd: bytes, wait: float = 0.3, timeout: float = 3.0) -> str:
        if self._usb_dev is not None:
            return self._cmd_usb(cmd, timeout=timeout)
        return self._cmd(cmd, wait=wait, timeout=timeout)

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

        # ── 표준 포맷: OK,Y,x,y,u',v',CCT,duv (또는 OK00,Y,x,y,...) ────
        if status.startswith("OK") and len(parts) >= 8:
            try:
                Lv  = float(parts[1])
                x   = float(parts[2])
                y   = float(parts[3])
                up  = float(parts[4])
                vp  = float(parts[5])
                cct = float(parts[6])
                duv = float(parts[7])
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
            if   raw1 <= 1.0:   cx, cy = raw1,           raw2
            elif raw1 < 10.0:   cx, cy = raw1 / 10.0,    raw2 / 10.0
            elif raw1 < 100.0:  cx, cy = raw1 / 100.0,   raw2 / 100.0
            elif raw1 < 1000.0: cx, cy = raw1 / 1000.0,  raw2 / 1000.0
            else:               cx, cy = raw1 / 10000.0, raw2 / 10000.0
            Lv = float(m.group(3))
            X, Y, Z = _xyz_from_Yxy(Lv, cx, cy)
            denom = -2 * cx + 12 * cy + 3
            up = (4 * cx / denom) if denom != 0 else 0.0
            vp = (9 * cy / denom) if denom != 0 else 0.0
            cct, duv = xy_to_cct_duv(cx, cy)
            return MeasureResult(
                timestamp_ms=timestamp_ms,
                Lv=Lv, x=cx, y=cy,
                u_prime=up, v_prime=vp,
                X=X, Y=Y, Z=Z,
                cct=cct, duv=duv,
                pattern_info=self._current_pattern,
            )

        raise RuntimeError(f"CA 미터 알 수 없는 응답 포맷: {resp!r}")

    # ------------------------------------------------------------------
    # Low-level serial I/O
    # ------------------------------------------------------------------

    def _cmd(self, cmd: bytes, wait: float = 0.1, timeout: float = 3.0) -> str:
        """시리얼 명령 전송 → 응답 한 줄 반환 (CA-410 / CA-310 serial 공용)."""
        assert self._serial is not None
        try:
            self._serial.reset_input_buffer()
            self._serial.write(cmd)
            self._serial.flush()
            if wait > 0:
                time.sleep(wait)

            self._serial.timeout = timeout
            raw = self._serial.read_until(b"\r")
            resp = raw.decode("ascii", errors="replace").strip()
            if resp:
                return resp

            time.sleep(0.05)
            extra = self._serial.read_all() or b""
            return extra.decode("ascii", errors="replace").strip()

        except Exception as e:
            raise RuntimeError(f"CA 미터 통신 오류: {e}") from e
