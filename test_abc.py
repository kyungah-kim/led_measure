#!/usr/bin/env python3
"""test_abc.py — SPTS4(0,1,2,6,28,N) 스캔: X(십자) 체크 코드 탐색.

확인된 상태:
  SPTS4(0,1,2,6,28) → ABC 버튼 + ㅁ 체크 + R+G+B 동시 활성
  목표: X(십자) 도 체크되는 N 코드 찾기

Usage:
    python test_abc.py /dev/ttyUSB0
"""
from __future__ import annotations

import sys
import time
import serial

_ACK = b'\x06'
_ETX = b'\x03'
_ENQ = b'\x05'
_SEP = b','
_PROG_4K60 = 2286


def _build(cmd1, cmd2, *params):
    body = _SEP.join(str(p).encode() for p in params)
    return b'\x02\xFD' + bytes([cmd1, cmd2]) + body + b'\x03'


def _send(ser, frame):
    ser.reset_input_buffer()
    ser.write(frame)
    deadline = time.time() + 1.5
    while time.time() < deadline:
        if ser.in_waiting:
            ch = ser.read(1)
            if ch == _ACK:
                return 'ok'
            if ch == b'\x11':
                time.sleep(0.2)
                ser.read(ser.in_waiting)
                return 'ests'
            if ch == _ETX:
                return 'data'
        time.sleep(0.01)
    return 'timeout'


def c(ser, frame, tag):
    r = _send(ser, frame)
    print(f"  {'✓' if r=='ok' else '✗'} {tag:<45s} → {r}")
    time.sleep(0.12)


def enter_tm(ser):
    for _ in range(3):
        ser.reset_input_buffer()
        ser.write(_ENQ)
        time.sleep(0.3)
        if ser.in_waiting and ser.read(1) == _ACK:
            return True
    return False


def load(ser):
    print("  4K60p 타이밍 로드…")
    c(ser, _build(0x24, 0x20, _PROG_4K60, 0), f"EXPDN4({_PROG_4K60},0)")
    time.sleep(2.5)


def confirm(ser, candidates: list[int]):
    """후보 코드 재확인: SPTS4(0,1,2,6,28,N) 각각 개별 테스트."""
    print(f"\n후보 코드 {candidates} 재확인")
    print("Enter=다음 / x=찾음! / q=종료\n")
    for n in candidates:
        c(ser, _build(0x20, 0x2C, 9999, 1,  255, 255, 255, 8), "SPT4 fg=white")
        c(ser, _build(0x20, 0x2C, 9999, 18, 0,   0,   0,   8), "SPT4 bg=black")
        c(ser, _build(0x20, 0x2A, 9999, 0, 1, 2, 6, 28, n),    f"SPTS4(0,1,2,6,28,{n})")
        c(ser, _build(0x24, 0x20, 9999, 0),                     "EXPDN4")
        ans = input(f"  [N={n}] X 체크됨? (Enter=다음 / x=이게맞음! / q=종료): ").strip().lower()
        if ans == 'x':
            print(f"\n★ X 코드 확정: {n} ★")
            print(f"  최종 조합: SPTS4(0,1,2,6,28,{n})")
            return n
        if ans == 'q':
            break
    return None


def scan(ser):
    """SPTS4(0,1,2,6,28,N) — N을 순서대로 바꿔가며 X 체크 코드 탐색."""
    skip = {0, 1, 2, 6, 28}
    print("\nSPTS4(0,1,2,6,28,N) 스캔 시작 — X 체크되는 N 찾기")
    print("각 코드마다 화면 확인 후 Enter(다음) / x(X체크됨!) / q(종료)\n")

    for n in range(0, 50):
        if n in skip:
            continue
        c(ser, _build(0x20, 0x2C, 9999, 1,  255, 255, 255, 8), "SPT4 fg=white")
        c(ser, _build(0x20, 0x2C, 9999, 18, 0,   0,   0,   8), "SPT4 bg=black")
        c(ser, _build(0x20, 0x2A, 9999, 0, 1, 2, 6, 28, n),    f"SPTS4(0,1,2,6,28,{n})")
        c(ser, _build(0x24, 0x20, 9999, 0),                     "EXPDN4")

        ans = input(f"  [N={n:2d}] ㅁ 유지? X 체크됨? (Enter=다음 / x=찾음! / q=종료): ").strip().lower()
        if ans == 'x':
            print(f"\n★ X 코드 발견: {n} ★")
            print(f"  최종 조합: SPTS4(0,1,2,6,28,{n})")
            break
        if ans == 'q':
            break


def main():
    port = sys.argv[1] if len(sys.argv) > 1 else "/dev/ttyUSB0"
    print(f"포트: {port}")

    ser = serial.Serial(port, 38400,
                        bytesize=serial.EIGHTBITS, parity=serial.PARITY_NONE,
                        stopbits=serial.STOPBITS_ONE, rtscts=False, timeout=1.0)
    try:
        ser.dtr, ser.rts = True, True
        time.sleep(0.5)
        if not enter_tm(ser):
            print("Terminal Mode 실패"); return
        print("✓ Terminal Mode")
        load(ser)
        confirm(ser, [25, 26])
    finally:
        ser.close()
        print("포트 닫힘.")


if __name__ == "__main__":
    main()


# ──────────────────────────────────────────────────────────────────────────────
# 아래는 수동 탐색용 케이스 메뉴 (필요 시 주석 해제 후 main() 에서 manual_menu(ser) 호출)
# ──────────────────────────────────────────────────────────────────────────────

# def fg(ser, r, g, b, bm=8):
#     c(ser, _build(0x20, 0x2C, 9999, 1, r, g, b, bm), f"SPT4 fg=({r},{g},{b})")
#
# def bg(ser, r=0, g=0, b=0, bm=8):
#     c(ser, _build(0x20, 0x2C, 9999, 18, r, g, b, bm), f"SPT4 bg=({r},{g},{b})")
#
# def spts4(ser, *codes):
#     c(ser, _build(0x20, 0x2A, 9999, *codes), f"SPTS4({','.join(map(str,codes))})")
#
# def expdn(ser):
#     c(ser, _build(0x24, 0x20, 9999, 0), "EXPDN4(9999,0)")
#     time.sleep(0.35)
#
# def step(ser, code, r, g, b):
#     """코드 하나 + EXPDN4 → 상태 누적 (ABC 활성화 방식)."""
#     fg(ser, r, g, b); bg(ser); spts4(ser, code); expdn(ser)
#
#
# CASES = {}
# def case(k, d):
#     def deco(fn): CASES[k] = (d, fn); return fn
#     return deco
#
#
# # ── 확인된 조합 ───────────────────────────────────────────────────────────────
#
# @case("A1", "step 0→1→2: ABC 버튼 + R+G+B 동시 (누적 방식)")
# def tA1(ser):
#     step(ser, 0, 255, 0, 0); step(ser, 1, 0, 255, 0); step(ser, 2, 0, 0, 255)
#     print("  → ABC 버튼 + R+G+B?")
#
# @case("A2", "SPTS4(0,1,2): R+G+B 동시 단일 호출")
# def tA2(ser):
#     fg(ser, 255, 255, 255); bg(ser); spts4(ser, 0, 1, 2); expdn(ser)
#     print("  → R+G+B 동시?")
#
# @case("B5a", "★ SPTS4(0,1,2,6): ABC 버튼 + ㅁ + R+G+B")
# def tB5a(ser):
#     fg(ser, 255, 255, 255); bg(ser); spts4(ser, 0, 1, 2, 6); expdn(ser)
#     print("  → ABC 버튼 + ㅁ 체크 + R+G+B?")
#
# @case("B5b", "★★ SPTS4(0,1,2,6,28): ABC + ㅁ + R+G+B (X 미체크)")
# def tB5b(ser):
#     fg(ser, 255, 255, 255); bg(ser); spts4(ser, 0, 1, 2, 6, 28); expdn(ser)
#     print("  → ABC + ㅁ 체크 + R+G+B (X 체크?)")
#
#
# # ── 수동 코드 입력 탐색 ───────────────────────────────────────────────────────
#
# @case("M1", "SPTS4(0,1,2,6,28,N) 수동 입력")
# def tM1(ser):
#     while True:
#         n = input("  추가 코드 (숫자 / q): ").strip()
#         if n.lower() == 'q': break
#         try:
#             code = int(n)
#             fg(ser, 255, 255, 255); bg(ser)
#             spts4(ser, 0, 1, 2, 6, 28, code); expdn(ser)
#             print(f"  → SPTS4(0,1,2,6,28,{code}): ABC? ㅁ? X?")
#         except ValueError:
#             print("  숫자를 입력하세요")
#
# @case("M2", "step 방식 수동: 0+E→1+E→2+E 후 코드 추가")
# def tM2(ser):
#     step(ser, 0, 255, 0, 0); step(ser, 1, 0, 255, 0); step(ser, 2, 0, 0, 255)
#     print("  [ABC + R+G+B 활성화됨]")
#     while True:
#         n = input("  추가 코드 (숫자 / q): ").strip()
#         if n.lower() == 'q': break
#         try:
#             code = int(n)
#             step(ser, code, 255, 255, 255)
#             print(f"  → step({code}): ABC 유지? ㅁ? X?")
#         except ValueError:
#             print("  숫자를 입력하세요")
#
#
# def manual_menu(ser):
#     print("\n케이스 목록:")
#     for k, (d, _) in CASES.items():
#         print(f"  {k} — {d}")
#     print("  q — 종료\n")
#     while True:
#         key = input("케이스: ").strip().upper()
#         if key == 'Q': break
#         elif key in CASES:
#             print(f"\n  실행: {key} — {CASES[key][0]}")
#             CASES[key][1](ser)
#         else:
#             print(f"  없는 케이스: {key}")
