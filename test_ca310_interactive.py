#!/usr/bin/env python3
"""test_ca310_interactive.py — CA-310 인터랙티브 연결 진단.

Usage:
    python test_ca310_interactive.py
    python test_ca310_interactive.py COM3
    python test_ca310_interactive.py /dev/ttyUSB0
"""
from __future__ import annotations

import sys
import time
import serial
import serial.tools.list_ports


# ── 유틸 ─────────────────────────────────────────────────────────────────────

def hex_str(data: bytes) -> str:
    return " ".join(f"{b:02X}" for b in data)


def list_ports() -> list[str]:
    ports = list(serial.tools.list_ports.comports())
    if not ports:
        print("  (연결된 COM 포트 없음)")
        return []
    for i, p in enumerate(ports):
        print(f"  {i+1}. {p.device:<15} {p.description}")
    return [p.device for p in ports]


def read_all(ser: serial.Serial, wait: float = 0.5) -> bytes:
    """wait 초 대기 후 버퍼에 있는 모든 바이트 읽기."""
    time.sleep(wait)
    data = b""
    ser.timeout = 0.3
    while True:
        chunk = ser.read(256)
        if not chunk:
            break
        data += chunk
    return data


def send_cmd(ser: serial.Serial, cmd: bytes, wait: float = 0.5) -> bytes:
    ser.reset_input_buffer()
    ser.write(cmd)
    ser.flush()
    return read_all(ser, wait)


def print_response(raw: bytes) -> None:
    if raw:
        txt = raw.decode("ascii", errors="replace")
        print(f"  ← 응답: {raw!r}")
        print(f"  ← hex:  [{hex_str(raw)}]")
        print(f"  ← text: {txt.strip()!r}")
    else:
        print("  ← (무응답)")


# ── 자동 스캔 ─────────────────────────────────────────────────────────────────

def auto_scan(port: str) -> None:
    """다양한 baud/흐름제어 조합을 시도해 응답이 오는 설정을 찾는다."""
    CONFIGS = [
        # (baud, bytesize, parity, stopbits, rtscts, dsrdtr, label)
        (115200, serial.EIGHTBITS, serial.PARITY_NONE, serial.STOPBITS_ONE, False, False, "115200 8N1"),
        (  9600, serial.EIGHTBITS, serial.PARITY_NONE, serial.STOPBITS_ONE, False, False,   "9600 8N1"),
        ( 19200, serial.EIGHTBITS, serial.PARITY_NONE, serial.STOPBITS_ONE, False, False,  "19200 8N1"),
        ( 38400, serial.EIGHTBITS, serial.PARITY_NONE, serial.STOPBITS_ONE, False, False,  "38400 8N1"),
        (115200, serial.EIGHTBITS, serial.PARITY_NONE, serial.STOPBITS_ONE, False,  True, "115200 8N1 DSR/DTR"),
        (  9600, serial.EIGHTBITS, serial.PARITY_NONE, serial.STOPBITS_ONE, False,  True,   "9600 8N1 DSR/DTR"),
        (115200, serial.SEVENBITS, serial.PARITY_EVEN,  serial.STOPBITS_TWO,  True, False, "115200 7E2 RTS/CTS"),
        (  9600, serial.SEVENBITS, serial.PARITY_EVEN,  serial.STOPBITS_ONE, False, False,   "9600 7E1"),
    ]
    CMDS = [b"COM,1\r\n", b"COM,1\r", b"IDQ\r\n", b"\r\n"]

    print(f"\n자동 스캔 시작 ({len(CONFIGS)} 설정 × {len(CMDS)} 명령)...\n")
    found = []

    for baud, bs, par, sb, rtscts, dsrdtr, label in CONFIGS:
        try:
            ser = serial.Serial(port, baudrate=baud, bytesize=bs, parity=par,
                                stopbits=sb, rtscts=rtscts, dsrdtr=dsrdtr, timeout=1.0)
        except Exception as e:
            print(f"  [{label}] 포트 열기 실패: {e}")
            continue

        try:
            ser.dtr = True
            ser.rts = False if not rtscts else ser.rts
            time.sleep(0.3)

            for cmd in CMDS:
                raw = send_cmd(ser, cmd, wait=0.4)
                if raw:
                    txt = raw.decode("ascii", errors="replace").strip()
                    print(f"  ✓ [{label}] cmd={cmd!r} → {raw!r}  text={txt!r}")
                    found.append((label, cmd, raw))
        finally:
            ser.close()
        time.sleep(0.1)

    if found:
        print(f"\n★ 응답 있는 설정 {len(found)}개:")
        for label, cmd, raw in found:
            print(f"   {label}  cmd={cmd!r}  응답={raw!r}")
    else:
        print("\n모든 설정에서 응답 없음.")
        print("확인: 케이블/전원/장치 메뉴(RS-232C 모드) / 장치관리자 포트번호")


# ── 인터랙티브 터미널 ──────────────────────────────────────────────────────────

def interactive(ser: serial.Serial, label: str) -> None:
    """열린 포트에 직접 명령을 입력하는 미니 터미널."""
    print(f"\n인터랙티브 모드 [{label}]")
    print("명령 입력 → Enter (자동으로 \\r\\n 추가)")
    print("특수 입력:")
    print("  :r    — CR(\\r) 만 전송")
    print("  :n    — LF(\\n) 만 전송")
    print("  :rn   — CR+LF 전송")
    print("  :hex AABB.. — 16진수 바이트 직접 전송")
    print("  :dtr0 / :dtr1 — DTR Low/High 전환")
    print("  :rts0 / :rts1 — RTS Low/High 전환")
    print("  :wait N — N초 대기 후 버퍼 읽기")
    print("  :q    — 종료\n")

    while True:
        try:
            line = input(">>> ").strip()
        except (EOFError, KeyboardInterrupt):
            break

        if not line:
            continue

        if line == ":q":
            break
        elif line == ":r":
            cmd = b"\r"
        elif line == ":n":
            cmd = b"\n"
        elif line == ":rn":
            cmd = b"\r\n"
        elif line.startswith(":hex "):
            try:
                cmd = bytes.fromhex(line[5:].replace(" ", ""))
            except ValueError:
                print("  16진수 형식 오류 (예: :hex 434F4D2C310D0A)")
                continue
        elif line == ":dtr0":
            ser.dtr = False
            print("  DTR → Low")
            continue
        elif line == ":dtr1":
            ser.dtr = True
            print("  DTR → High")
            continue
        elif line == ":rts0":
            ser.rts = False
            print("  RTS → Low")
            continue
        elif line == ":rts1":
            ser.rts = True
            print("  RTS → High")
            continue
        elif line.startswith(":wait "):
            try:
                wait = float(line[6:])
                raw = read_all(ser, wait)
                print(f"  {wait}초 대기 후 읽기:")
                print_response(raw)
            except ValueError:
                print("  숫자 입력 필요 (예: :wait 1.5)")
            continue
        else:
            cmd = (line + "\r\n").encode("ascii")

        print(f"  → 전송: {cmd!r}  hex=[{hex_str(cmd)}]")
        raw = send_cmd(ser, cmd, wait=0.5)
        print_response(raw)


# ── 설정 선택 메뉴 ────────────────────────────────────────────────────────────

PRESET_CONFIGS = [
    ("115200 8N1 (기본)", dict(baudrate=115200, bytesize=serial.EIGHTBITS,
                               parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                               rtscts=False, dsrdtr=False)),
    ("9600 8N1",          dict(baudrate=9600,   bytesize=serial.EIGHTBITS,
                               parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                               rtscts=False, dsrdtr=False)),
    ("19200 8N1",         dict(baudrate=19200,  bytesize=serial.EIGHTBITS,
                               parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                               rtscts=False, dsrdtr=False)),
    ("38400 8N1",         dict(baudrate=38400,  bytesize=serial.EIGHTBITS,
                               parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                               rtscts=False, dsrdtr=False)),
    ("9600 8N1 + DSR/DTR",dict(baudrate=9600,   bytesize=serial.EIGHTBITS,
                               parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                               rtscts=False, dsrdtr=True)),
    ("115200 8N1 + DSR/DTR",dict(baudrate=115200,bytesize=serial.EIGHTBITS,
                               parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                               rtscts=False, dsrdtr=True)),
    ("115200 7E2 + RTS/CTS (CA-410 방식)", dict(baudrate=115200, bytesize=serial.SEVENBITS,
                               parity=serial.PARITY_EVEN, stopbits=serial.STOPBITS_TWO,
                               rtscts=True, dsrdtr=False)),
    ("9600 7E1",          dict(baudrate=9600,   bytesize=serial.SEVENBITS,
                               parity=serial.PARITY_EVEN, stopbits=serial.STOPBITS_ONE,
                               rtscts=False, dsrdtr=False)),
]


def select_config() -> tuple[str, dict] | None:
    print("\n설정 선택:")
    for i, (label, _) in enumerate(PRESET_CONFIGS):
        print(f"  {i+1}. {label}")
    print("  0. 취소")
    sel = input("선택: ").strip()
    try:
        idx = int(sel)
        if idx == 0:
            return None
        label, cfg = PRESET_CONFIGS[idx - 1]
        return label, cfg
    except (ValueError, IndexError):
        print("잘못된 입력")
        return None


# ── main ─────────────────────────────────────────────────────────────────────

def main() -> None:
    print("=" * 55)
    print("CA-310 인터랙티브 연결 진단")
    print("=" * 55)

    # 포트 선택
    if len(sys.argv) >= 2:
        port = sys.argv[1]
    else:
        print("\n사용 가능한 COM 포트:")
        ports = list_ports()
        if not ports:
            return
        sel = input("\n포트 번호 선택 (1~) 또는 직접 입력: ").strip()
        try:
            port = ports[int(sel) - 1]
        except (ValueError, IndexError):
            port = sel

    print(f"\n선택된 포트: {port}")

    while True:
        print("\n" + "─" * 40)
        print("메뉴")
        print("  1. 자동 스캔 (모든 설정 시도)")
        print("  2. 설정 선택 후 인터랙티브 터미널")
        print("  3. 포트 변경")
        print("  0. 종료")
        menu = input("선택: ").strip()

        if menu == "0":
            break

        elif menu == "1":
            auto_scan(port)

        elif menu == "2":
            result = select_config()
            if result is None:
                continue
            label, cfg = result
            print(f"\n  설정: {label}")
            try:
                ser = serial.Serial(port, timeout=1.0, **cfg)
                try:
                    ser.dtr = True
                    if not cfg.get("rtscts"):
                        ser.rts = False
                    time.sleep(0.3)
                    print(f"  ✓ 포트 열림  DTR={ser.dtr}  RTS={ser.rts}  DSR={ser.dsr}  CTS={ser.cts}")
                    interactive(ser, label)
                finally:
                    ser.close()
                    print("  포트 닫힘")
            except Exception as e:
                print(f"  ✗ 포트 열기 실패: {e}")

        elif menu == "3":
            print("\n사용 가능한 COM 포트:")
            ports = list_ports()
            sel = input("포트 번호 선택 또는 직접 입력: ").strip()
            try:
                port = ports[int(sel) - 1]
            except (ValueError, IndexError):
                port = sel
            print(f"포트 변경: {port}")

    print("\n종료.")


if __name__ == "__main__":
    main()
