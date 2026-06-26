#!/usr/bin/env python3
"""test_ca310.py — CA-310 시리얼 연결 진단 도구.

Usage:
    python test_ca310.py COM3          # Windows
    python test_ca310.py /dev/ttyUSB0  # Linux
    python test_ca310.py               # COM 포트 목록만 출력
"""
from __future__ import annotations

import sys
import time
import serial
import serial.tools.list_ports


# ── 진단 설정 ────────────────────────────────────────────────────────────────

SCAN_BAUD   = [9600, 19200, 38400, 57600, 115200]
SCAN_CMDS   = [b"COM,1\r\n", b"COM,1\r", b"COM,1\n", b"\r", b"IDQ\r\n"]
LISTEN_SEC  = 1.0   # 각 baud 에서 자발적 출력 대기 시간


def list_ports() -> None:
    ports = list(serial.tools.list_ports.comports())
    if not ports:
        print("연결된 COM 포트 없음")
        return
    print(f"{'포트':<15} {'설명'}")
    for p in ports:
        print(f"{p.device:<15} {p.description}")


def raw_hex(data: bytes) -> str:
    return " ".join(f"{b:02X}" for b in data)


def try_port(port: str) -> None:
    print(f"\n{'='*60}")
    print(f"포트: {port}")
    print(f"{'='*60}")

    for baud in SCAN_BAUD:
        print(f"\n  [baud={baud}]")
        try:
            ser = serial.Serial(
                port,
                baudrate=baud,
                bytesize=serial.EIGHTBITS,
                parity=serial.PARITY_NONE,
                stopbits=serial.STOPBITS_ONE,
                rtscts=False,
                dsrdtr=False,
                timeout=1.0,
            )
        except Exception as e:
            print(f"    ✗ 포트 열기 실패: {e}")
            continue

        try:
            # DTR High 설정 (CA-310 응답에 필요)
            ser.dtr = True
            ser.rts = False
            time.sleep(0.3)

            # ── 자발적 출력 대기 (장치가 먼저 뭔가 보내는지) ────────────
            ser.reset_input_buffer()
            time.sleep(LISTEN_SEC)
            passv = ser.read(ser.in_waiting or 0)
            if passv:
                print(f"    자발적 출력: {passv!r}  hex=[{raw_hex(passv)}]")
            else:
                print(f"    자발적 출력: (없음)")

            # ── 각 명령 시도 ─────────────────────────────────────────────
            for cmd in SCAN_CMDS:
                ser.reset_input_buffer()
                ser.write(cmd)
                ser.flush()
                time.sleep(0.5)

                raw = ser.read(ser.in_waiting or 1)  # 있는 것 모두 읽기
                if not raw:
                    # read_until 으로 한 번 더 시도 (최대 0.5초)
                    ser.timeout = 0.5
                    raw = ser.read_until(b"\r")

                if raw:
                    txt = raw.decode("ascii", errors="replace").strip()
                    print(f"    cmd={cmd!r:20s} → {raw!r}  txt={txt!r}")
                else:
                    print(f"    cmd={cmd!r:20s} → (무응답)")

        finally:
            ser.close()

        # 첫 번째로 응답이 온 baud 를 찾으면 중단할 수도 있지만
        # 진단용이므로 전부 출력
        print()


def main() -> None:
    if len(sys.argv) < 2:
        print("사용법: python test_ca310.py <PORT>")
        print()
        list_ports()
        return

    port = sys.argv[1]
    print("CA-310 연결 진단 시작...")
    print("각 baud rate 에서 5가지 명령을 전송하고 응답을 출력합니다.\n")

    try_port(port)

    print("\n진단 완료.")
    print("응답이 있는 baud rate 와 명령을 확인해 core/equipment/ca_meter.py 에 반영하세요.")


if __name__ == "__main__":
    main()
