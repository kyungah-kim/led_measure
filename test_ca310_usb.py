#!/usr/bin/env python3
"""test_ca310_usb.py — CA-310 USB 프로토콜 탐색.

Usage:
    sudo python test_ca310_usb.py
"""
from __future__ import annotations
import sys
import time
from typing import Any

try:
    import usb.core
    import usb.util
except ImportError:
    print("pyusb 없음: pip install pyusb")
    sys.exit(1)

VID = 0x0686
PID = 0x1002

# 엔드포인트 쌍
EP_PAIRS = [
    (0x01, 0x81, "EP1 OUT→IN (16-byte)"),
    (0x02, 0x82, "EP2 OUT→IN (64-byte)"),
    (0x01, 0x82, "EP1 OUT, EP2 IN (cross)"),
    (0x02, 0x81, "EP2 OUT, EP1 IN (cross)"),
]

COMMANDS = [
    b"COM,1\r\n",
    b"COM,1\r",
    b"IDQ\r\n",
    b"\r\n",
    b"\r",
]


def try_bulk(dev, ep_out: int, ep_in: int, cmd: bytes, label: str, timeout_ms: int = 2000) -> None:
    # 일부 장치는 명령을 패킷 크기에 맞게 패딩 요구
    for pad in [False, True]:
        data = cmd
        if pad:
            # 16바이트 또는 64바이트로 패딩
            pkt_size = 64 if ep_out == 0x02 else 16
            if len(data) < pkt_size:
                data = data + b"\x00" * (pkt_size - len(data))
        tag = f"  [{label}] {'pad' if pad else 'raw'} cmd={cmd!r}"
        try:
            dev.write(ep_out, data, timeout=timeout_ms)
            time.sleep(0.3)
            raw = bytes(dev.read(ep_in, 512, timeout=timeout_ms))
            resp = raw.decode("ascii", errors="replace").strip()
            print(f"{tag}")
            print(f"    수신: {raw!r}  text={resp!r}")
            return  # 응답 있으면 패딩 반복 중단
        except usb.core.USBTimeoutError:
            print(f"{tag}  → 타임아웃")
        except Exception as e:
            print(f"{tag}  → 오류: {e}")


def try_passive_listen(dev, ep_in: int, label: str, duration: float = 1.5) -> None:
    """명령 없이 수신만 — 장치가 자발적으로 뭔가 보내는지 확인."""
    print(f"  [수동 대기 {label}] {duration}초 대기...")
    deadline = time.time() + duration
    while time.time() < deadline:
        try:
            raw = bytes(dev.read(ep_in, 512, timeout=200))
            if raw:
                resp = raw.decode("ascii", errors="replace").strip()
                print(f"    자발적 수신: {raw!r}  text={resp!r}")
                return
        except usb.core.USBTimeoutError:
            pass
        except Exception as e:
            print(f"    오류: {e}")
            return
    print("    (자발적 수신 없음)")


def try_control(dev) -> None:
    """Control transfer — vendor-specific 초기화 시도."""
    print("\n[Control Transfer 시도]")
    # bmRequestType=0x40: host→device, vendor, device
    # bRequest 값은 벤더에 따라 다름 — 0x01~0x10 범위 시도
    for req in [0x01, 0x02, 0x09, 0x0A]:
        try:
            ret = dev.ctrl_transfer(0x40, req, 0, 0, None, timeout=1000)
            print(f"  ctrl OUT req=0x{req:02X} → 성공 ({ret} bytes sent)")
            # 응답 읽기
            try:
                raw = bytes(dev.ctrl_transfer(0xC0, req, 0, 0, 64, timeout=1000))
                print(f"  ctrl IN  req=0x{req:02X} → {raw!r}")
            except Exception:
                pass
        except Exception as e:
            print(f"  ctrl req=0x{req:02X} → {e}")


EP_OUT = 0x02
EP_IN  = 0x82


def send(dev: Any, cmd: bytes, timeout_ms: int = 3000) -> str:
    dev.write(EP_OUT, cmd, timeout=timeout_ms)
    time.sleep(0.3)
    raw = bytes(dev.read(EP_IN, 512, timeout=timeout_ms))
    resp = raw.decode("ascii", errors="replace").strip()
    print(f"  → {cmd!r}")
    print(f"  ← {raw[:40]!r}  text={resp!r}")
    return resp


def find_dev() -> Any:
    dev: Any = usb.core.find(idVendor=VID, idProduct=PID)  # type: ignore[assignment]
    return dev


def open_dev(dev: Any) -> None:
    try:
        if dev.is_kernel_driver_active(0):
            dev.detach_kernel_driver(0)
    except Exception:
        pass
    dev.set_configuration()


def read_both(dev: Any, timeout_ms: int = 1000, label: str = "") -> dict[str, bytes | None]:
    result: dict[str, bytes | None] = {}
    for ep, name in [(0x81, "EP1 IN"), (0x82, "EP2 IN")]:
        try:
            raw = bytes(dev.read(ep, 512, timeout=timeout_ms))
            print(f"  [{label}] {name}: {raw[:40]!r}  text={raw.decode('ascii','replace').strip()!r}")
            result[name] = raw
        except usb.core.USBTimeoutError:
            print(f"  [{label}] {name}: (없음)")
            result[name] = None
    return result


def main() -> None:
    dev = find_dev()
    if dev is None:
        print("CA-310 USB 장치를 찾을 수 없음")
        return
    print(f"CA-310 발견  Bus {dev.bus} Device {dev.address}\n")

    # ── A. USB RESET → 재탐색 → set_configuration ─────────────────
    print("=== A. USB reset 후 재탐색 ===")
    try:
        dev.reset()
        print("  reset() 완료")
    except Exception as e:
        print(f"  reset() 실패: {e}")
    time.sleep(1.5)

    dev = find_dev()
    if dev is None:
        print("  reset 후 장치 재탐색 실패 — set_configuration 생략")
    else:
        print(f"  재탐색  Bus {dev.bus} Device {dev.address}")
        open_dev(dev)
        time.sleep(1.0)

        print("  자발적 응답 대기 (3초)...")
        deadline = time.time() + 3.0
        while time.time() < deadline:
            for ep, name in [(0x81, "EP1 IN"), (0x82, "EP2 IN")]:
                try:
                    raw = bytes(dev.read(ep, 512, timeout=200))
                    print(f"    {name} 자발적: {raw[:40]!r}")
                except usb.core.USBTimeoutError:
                    pass
        print("  (3초 경과)")

    # ── B. COM,1 — 64바이트 패딩 ──────────────────────────────────
    print("\n=== B. COM,1 64바이트 패딩 ===")
    padded = b"COM,1\r\n" + b"\x00" * (64 - len(b"COM,1\r\n"))
    try:
        dev.write(0x02, padded, timeout=2000)
        print("  EP2 OUT (64-pad) 성공")
    except Exception as e:
        print(f"  EP2 OUT (64-pad) 실패: {e}")
    read_both(dev, timeout_ms=3000, label="B")

    # ── C. COM,1 반복 전송 (최대 5회, 응답 올 때까지) ───────────────
    print("\n=== C. COM,1 반복 전송 ===")
    for i in range(5):
        try:
            dev.write(0x02, b"COM,1\r\n", timeout=2000)
        except Exception as e:
            print(f"  #{i+1} write 실패: {e}")
            continue
        time.sleep(0.2)
        # EP2 IN — 짧은 타임아웃으로 시도
        for ep, name in [(0x82, "EP2 IN"), (0x81, "EP1 IN")]:
            try:
                raw = bytes(dev.read(ep, 512, timeout=1500))
                resp = raw.decode("ascii", errors="replace").strip()
                print(f"  #{i+1} {name} 응답: {raw[:40]!r}  text={resp!r}")
                break
            except usb.core.USBTimeoutError:
                print(f"  #{i+1} {name}: (없음)")
        time.sleep(0.3)

    # ── D. 전체 측정 시퀀스 (64바이트 패딩 적용) ──────────────────────
    print("\n=== D. 전체 측정 시퀀스 ===")

    def send64(dev: Any, cmd: bytes, label: str = "") -> str:
        padded = cmd + b"\x00" * max(0, 64 - len(cmd))
        try:
            dev.write(0x02, padded, timeout=3000)
        except Exception as e:
            print(f"  [{label or cmd!r}] write 실패: {e}")
            return ""
        time.sleep(0.2)
        try:
            raw = bytes(dev.read(0x82, 512, timeout=5000))
            resp = raw.decode("ascii", errors="replace").strip()
            print(f"  [{label or cmd!r}] → {raw[:60]!r}  text={resp!r}")
            return resp
        except usb.core.USBTimeoutError:
            print(f"  [{label or cmd!r}] → (타임아웃)")
            return ""

    send64(dev, b"COM,1\r\n", "COM,1")
    send64(dev, b"MDS,0\r\n", "MDS,0")   # xyLv 모드
    print("  ★ MES 전송 (측정 중...):")
    send64(dev, b"MES\r\n", "MES")
    send64(dev, b"COM,0\r\n", "COM,0")   # REMOTE 해제

    usb.util.dispose_resources(dev)
    print("\n완료")


if __name__ == "__main__":
    main()
