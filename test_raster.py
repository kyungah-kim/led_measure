#!/usr/bin/env python3
"""test_raster.py — VG-879 White Raster + Black Window 명암비 패턴 테스트.

확인된 사항:
  SPTS4(0, 1, 2, 10) → White Raster (Raster+R+G+B 동시 활성)

다음 목표:
  White Raster 위에 중앙 Black Window 를 APL 순서대로 출력.
  APL 100% → 흰 래스터만
  APL  50% → 흰 배경 + 검은 창 side=70.7%
  APL  20% → 흰 배경 + 검은 창 side=89.4%
  APL  14.1% → 흰 배경 + 검은 창 side=92.7%

Usage:
    python test_raster.py /dev/ttyUSB0
    python test_raster.py /dev/ttyUSB0 --scan   # Raster 코드 재탐색
"""
from __future__ import annotations

import sys
import time
import serial

# ── 프로토콜 상수 ─────────────────────────────────────────────────────────────
_ACK       = b'\x06'
_ETX       = b'\x03'
_ENQ       = b'\x05'
_SEP       = b','
_PROG_4K60 = 2286
_SCREEN_W  = 3840
_SCREEN_H  = 2160

# 확인된 조합
RASTER_COMBO = (0, 1, 2, 10)   # SPTS4 → White Raster (Raster+R+G+B 동시)

# 검은 창 H/V 비율 스텝 (% of screen side)
_WIN_SIDES = [100.0, 50.0, 20.0, 14.1]


# ── 저수준 통신 ───────────────────────────────────────────────────────────────

def _build(cmd1: int, cmd2: int, *params) -> bytes:
    body = _SEP.join(str(p).encode() for p in params)
    return b'\x02\xFD' + bytes([cmd1, cmd2]) + body + b'\x03'


def _send(ser: serial.Serial, frame: bytes, timeout: float = 1.5) -> str:
    ser.reset_input_buffer()
    ser.write(frame)
    deadline = time.time() + timeout
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


def c(ser, frame, tag, sleep=0.12):
    r = _send(ser, frame)
    print(f"  {'✓' if r=='ok' else '✗'} {tag:<55s} → {r}")
    time.sleep(sleep)
    return r


def enter_terminal(ser) -> bool:
    for _ in range(3):
        ser.reset_input_buffer()
        ser.write(_ENQ)
        time.sleep(0.3)
        if ser.in_waiting and ser.read(1) == _ACK:
            return True
    return False


def load_timing(ser) -> None:
    c(ser, _build(0x24, 0x20, _PROG_4K60, 0), "EXPDN4(2286,0): 타이밍 로드", sleep=1.5)


# ── 패턴 헬퍼 ────────────────────────────────────────────────────────────────

def _center_coords(side_pct: float):
    w = max(1, int(_SCREEN_W * side_pct / 100))
    h = max(1, int(_SCREEN_H * side_pct / 100))
    x1 = (_SCREEN_W - w) // 2
    y1 = (_SCREEN_H - h) // 2
    return x1, y1, x1 + w - 1, y1 + h - 1


def reset(ser) -> None:
    """ALLCLR4: 창 플레인 + 베이스 전체 초기화."""
    c(ser, _build(0x28, 0x60), "ALLCLR4: 전체 초기화")


def white_raster(ser) -> None:
    """ALLCLR4 후 SPTS4(0,1,2,10) — White Raster 출력."""
    reset(ser)
    c(ser, _build(0x20, 0x2C, 9999,  1, 255, 255, 255, 8), "SPT4 fg=white")
    c(ser, _build(0x20, 0x2C, 9999, 18,   0,   0,   0, 8), "SPT4 bg=black")
    c(ser, _build(0x20, 0x2A, 9999, *RASTER_COMBO),         f"SPTS4{RASTER_COMBO}: White Raster")
    c(ser, _build(0x24, 0x20, 9999, 0),                     "EXPDN4(9999,0)")


def black_window_on_raster(ser, side_pct: float) -> None:
    """White Raster 위에 중앙 Black Window 추가."""
    x1, y1, x2, y2 = _center_coords(side_pct)
    c(ser, _build(0x28, 0x61, x1, y1, x2, y2),  f"WINDOW4: center side={side_pct:.1f}%")
    c(ser, _build(0x28, 0x62,   0,   0,   0, 8), "WINCOL4(black)")
    c(ser, _build(0x24, 0x20,   0,   0),          "EXPDN4(0,0)")


# ── 명암비 순차 테스트 ────────────────────────────────────────────────────────

def test_contrast_sequence(ser) -> None:
    """White Raster 한 번 출력 후 WINCLR4로 창만 교체하며 순차 테스트.

    흰 래스터는 처음 한 번만 설정.
    각 크기 전환 시 WINCLR4(창 플레인만 클리어)로 이전 창 제거 후 새 창 등록.
    ALLCLR4는 래스터까지 날리므로 사용 안 함.
    """
    print("\n" + "="*60)
    print(" White Raster + Black Window 순차 테스트")
    print(" 흰 래스터 1회 설정 → WINCLR4 + 창 크기 교체")
    print(" 창 H/V: 100% → 50% → 20% → 14.1%")
    print("="*60)

    # ── 흰 래스터 최초 1회 설정 ──────────────────────────────────────────────
    print("\n  [Step 0] White Raster 설정 (1회)")
    white_raster(ser)
    time.sleep(0.3)
    ans = input("  → 전체 흰색 확인? [Enter=계속 / q=종료]: ").strip().lower()
    if ans == 'q':
        return

    # ── 각 창 크기 순차 출력 ─────────────────────────────────────────────────
    for i, side in enumerate(_WIN_SIDES):
        print(f"\n{'─'*60}")
        print(f"  [Step {i+1}] 창 H/V {side:.1f}% × {side:.1f}%")
        print(f"{'─'*60}")

        # WINCLR4: 이전 창 플레인만 클리어 (래스터 베이스 유지 기대)
        c(ser, _build(0x28, 0x63), "WINCLR4: 이전 창 클리어")
        time.sleep(0.1)

        black_window_on_raster(ser, side)
        time.sleep(0.3)

        ans = input(
            f"  창 {side:.1f}%  결과?\n"
            f"  [1] 흰 배경 유지 + 검은 박스  ✓\n"
            f"  [2] 흰 배경 사라짐 (래스터 날아감)\n"
            f"  [3] 전체 검정\n"
            f"  [Enter] 다음  /  [q] 종료: "
        ).strip().lower()

        if ans == "2":
            print("\n  WINCLR4가 래스터도 지움 → 매 스텝 래스터 재출력 필요.")
            cont = input("  매 스텝 white_raster() 재출력 방식으로 재시도? [y/n]: ").strip().lower()
            if cont == 'y':
                _test_with_raster_each_step(ser)
            break
        elif ans == "3":
            print("\n  전체 검정 → WINCOL4가 래스터 덮음")
            break
        elif ans == "q":
            break
        elif ans == "1":
            print(f"  ✓ 창 {side:.1f}% 정상")

    print("\n  순차 테스트 완료")


def _test_with_raster_each_step(ser) -> None:
    """매 스텝 white_raster() 재출력 방식 (ALLCLR4 포함)."""
    print("\n  [매 스텝 래스터 재출력 방식]")
    for side in _WIN_SIDES:
        print(f"\n  창 H/V {side:.1f}%")
        white_raster(ser)   # ALLCLR4 + SPTS4 래스터
        time.sleep(0.1)
        black_window_on_raster(ser, side)
        time.sleep(0.3)
        ans = input(f"  창 {side:.1f}%: [1]=정상 / [Enter]=다음 / [q]=종료: ").strip().lower()
        if ans == 'q':
            break
        if ans == '1':
            print(f"  ✓ 창 {side:.1f}% 정상")


def _try_winclr(ser, side_pct: float) -> None:
    """WINCLR4로 창 플레인만 초기화 후 Black Window 재시도."""
    print(f"\n  [WINCLR4 변형]  side={side_pct:.1f}%")
    white_raster(ser)
    time.sleep(0.15)
    c(ser, _build(0x28, 0x63), "WINCLR4: 창 플레인만 클리어")
    time.sleep(0.15)
    black_window_on_raster(ser, side_pct)
    ans = input("  결과? [1]=흰배경+검은창  [2]=전체검정  [3]=기타: ").strip()
    if ans == "1":
        print("\n  ★ WINCLR4 방식 성공!")
        print("  → white_raster() → WINCLR4 → WINDOW4(black) + WINCOL4(black)")
    else:
        print(f"  결과: {ans}")


# ── Raster 코드 스캔 (필요 시) ────────────────────────────────────────────────

def scan_raster(ser) -> None:
    """SPTS4(N) 스캔으로 Raster 버튼 활성 코드 탐색 (N=0..59)."""
    print("\n스캔: SPTS4(N)  N=0..59  |  Raster 버튼 켜지면 r")
    for n in range(60):
        c(ser, _build(0x20, 0x2C, 9999,  1, 255, 255, 255, 8), "SPT4 fg=white")
        c(ser, _build(0x20, 0x2C, 9999, 18,   0,   0,   0, 8), "SPT4 bg=black")
        c(ser, _build(0x20, 0x2A, 9999, n),                     f"SPTS4({n})")
        c(ser, _build(0x24, 0x20, 9999, 0),                     "EXPDN4(9999,0)")
        ans = input(f"  [N={n:2d}] Raster 켜짐? (Enter=다음 / r=찾음 / q=종료): ").strip().lower()
        if ans == 'r':
            print(f"\n  ★ Raster 코드: N={n}")
            break
        if ans == 'q':
            break


def scan_raster_combo(ser) -> None:
    """SPTS4(0,1,2,10) 등 조합으로 흰색 래스터 출력 확인."""
    combos = [
        (0, 1, 2, 10),
        (10, 0, 1, 2),
        (0, 1, 2, 6, 10),
        (0, 1, 2, 6, 28, 10),
    ]
    print("\n  ※ 단일 SPTS4 호출 — 각 조합은 독립적")
    for combo in combos:
        print(f"\n  전송: SPTS4{combo}")
        c(ser, _build(0x20, 0x2C, 9999,  1, 255, 255, 255, 8), "SPT4 fg=white")
        c(ser, _build(0x20, 0x2C, 9999, 18,   0,   0,   0, 8), "SPT4 bg=black")
        c(ser, _build(0x20, 0x2A, 9999, *combo),                f"SPTS4{combo}")
        c(ser, _build(0x24, 0x20, 9999, 0),                     "EXPDN4(9999,0)")
        ans = input("  → 흰색+Raster+RGB 동시? [y=성공 / Enter=다음 / q=종료]: ").strip().lower()
        if ans == 'y':
            print(f"\n  ★ 성공! SPTS4{combo}")
            return
        if ans == 'q':
            break


# ── 메인 ─────────────────────────────────────────────────────────────────────

def main() -> None:
    port = "/dev/ttyUSB0"
    do_scan = False

    for arg in sys.argv[1:]:
        if arg == "--scan":
            do_scan = True
        else:
            port = arg

    print(f"포트: {port}")
    ser = serial.Serial(
        port, baudrate=38400, bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
        rtscts=False, timeout=1.0,
    )
    time.sleep(0.3)

    print("터미널 모드 진입...")
    if not enter_terminal(ser):
        print("실패 — 장비/케이블 확인")
        ser.close()
        sys.exit(1)
    print("OK")

    load_timing(ser)

    if do_scan:
        scan_raster(ser)
    else:
        # SPTS4(0,1,2,10) = White Raster 확인됨 → 바로 순차 테스트
        print(f"\n확인된 White Raster: SPTS4{RASTER_COMBO}")
        test_contrast_sequence(ser)

    ser.close()
    print("종료")


if __name__ == "__main__":
    main()
