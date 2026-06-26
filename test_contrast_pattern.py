#!/usr/bin/env python3
"""test_contrast_pattern.py — VG-879 명암비 패턴 방식 확인 테스트.

목표: White Raster + Black Window 패턴을 VG-879에서 구현하는 올바른 방법 탐색.

방식 A (현재 코드): 4-스트립 White Border + 미등록 Black Center
  - ALLCLR4 → 상/하/좌/우 4개 WINDOW4(흰 테두리) → WINCOL4(white) → EXPDN4
  - 문제: VG-879 WINDOW4 슬롯 한계(2개)로 좌/우가 렌더링 안 됨
    → 결과: 상/하 흰 스트립 + 중앙 전체 검은 수평밴드 (박스 아님)

방식 B (BCOL4): Background Color + Black Center Window
  - ALLCLR4 → BCOL4(white) → WINDOW4(center) → WINCOL4(black) → EXPDN4
  - BCOL4가 ALLCLR4 이후에 동작하면 흰 배경 + 검은 창이 가능

방식 C (Full-field White + WINCLR4 + Black Window):
  - full_field(white): ALLCLR4 → WINDOW4(full) → WINCOL4(white) → EXPDN4
  - 이후 WINCLR4(window 플레인만 초기화) → WINDOW4(center) → WINCOL4(black) → EXPDN4
  - 이론: full-field white를 베이스로 유지하면서 중앙만 검은 창 등록

각 방식을 실행 후 화면을 직접 확인해 어느 방식이 올바른 패턴을 출력하는지 확인.

Usage:
    python test_contrast_pattern.py /dev/ttyUSB0
"""
from __future__ import annotations

import sys
import time
import serial

# ── VG-879 프로토콜 상수 ──────────────────────────────────────────────────────
_ACK  = b'\x06'
_ETX  = b'\x03'
_ENQ  = b'\x05'
_SEP  = b','
_PROG_4K60 = 2286
_SCREEN_W  = 3840
_SCREEN_H  = 2160


def _build(cmd1: int, cmd2: int, *params) -> bytes:
    body = _SEP.join(str(p).encode() for p in params)
    return b'\x02\xFD' + bytes([cmd1, cmd2]) + body + b'\x03'


def _send(ser: serial.Serial, frame: bytes, timeout: float = 2.0) -> str:
    ser.reset_input_buffer()
    ser.write(frame)
    deadline = time.time() + timeout
    while time.time() < deadline:
        if ser.in_waiting:
            ch = ser.read(1)
            if ch == _ACK:
                return 'ok'
            if ch == b'\x11':          # ESTS
                time.sleep(0.2)
                ser.read(ser.in_waiting)
                return 'ests'
            if ch == _ETX:
                return 'data'
        time.sleep(0.01)
    return 'timeout'


def c(ser: serial.Serial, frame: bytes, tag: str, sleep: float = 0.1) -> str:
    r = _send(ser, frame)
    mark = '✓' if r == 'ok' else '✗'
    print(f"  {mark} {tag:<55s} → {r}")
    time.sleep(sleep)
    return r


def enter_terminal(ser: serial.Serial) -> bool:
    for _ in range(3):
        ser.reset_input_buffer()
        ser.write(_ENQ)
        time.sleep(0.3)
        if ser.in_waiting and ser.read(1) == _ACK:
            print("  터미널 모드 진입 OK")
            return True
    print("  터미널 모드 진입 실패")
    return False


def load_timing(ser: serial.Serial) -> None:
    print("  4K60p 타이밍 로드 (EXPDN4 2286)...")
    c(ser, _build(0x24, 0x20, _PROG_4K60, 0), "EXPDN4(2286,0): 타이밍 로드", sleep=1.5)


# ── 창 좌표 계산 ──────────────────────────────────────────────────────────────

def center_window_coords(win_pct: float):
    """중앙 정렬 창의 픽셀 좌표 반환 (x1, y1, x2, y2)."""
    w = max(1, int(_SCREEN_W * win_pct / 100))
    h = max(1, int(_SCREEN_H * win_pct / 100))
    x1 = (_SCREEN_W - w) // 2
    y1 = (_SCREEN_H - h) // 2
    return x1, y1, x1 + w - 1, y1 + h - 1


# ── 방식 A: 4-스트립 흰 테두리 (현재 구현) ───────────────────────────────────

def method_a_four_strips(ser: serial.Serial, win_pct: float) -> None:
    """현재 코드의 bright-bg 방식: 4-strip white border + unregistered black center."""
    x1, y1, x2, y2 = center_window_coords(win_pct)
    print(f"\n  창 좌표: ({x1},{y1}) ~ ({x2},{y2})  창 크기: {win_pct:.1f}%")

    c(ser, _build(0x28, 0x60), "ALLCLR4: 전체 클리어")
    # 상단
    if y1 > 0:
        c(ser, _build(0x28, 0x61, 0, 0, _SCREEN_W-1, y1-1), f"WINDOW4: 상단 스트립 (0→{y1-1})")
    # 하단
    if y2 < _SCREEN_H - 1:
        c(ser, _build(0x28, 0x61, 0, y2+1, _SCREEN_W-1, _SCREEN_H-1), f"WINDOW4: 하단 스트립 ({y2+1}→H)")
    # 좌측 (슬롯 부족 시 실패)
    if x1 > 0:
        c(ser, _build(0x28, 0x61, 0, y1, x1-1, y2), f"WINDOW4: 좌측 스트립 (0→{x1-1})")
    # 우측 (슬롯 부족 시 실패)
    if x2 < _SCREEN_W - 1:
        c(ser, _build(0x28, 0x61, x2+1, y1, _SCREEN_W-1, y2), f"WINDOW4: 우측 스트립 ({x2+1}→W)")

    c(ser, _build(0x28, 0x62, 255, 255, 255, 8), "WINCOL4(white): 스트립 색상")
    c(ser, _build(0x24, 0x20, 0, 0), "EXPDN4(0,0): 렌더링")


# ── 방식 B: BCOL4 배경색 설정 + Black Center Window ──────────────────────────

def method_b_bcol4(ser: serial.Serial, win_pct: float) -> None:
    """BCOL4로 배경을 흰색으로, WINDOW4+WINCOL4로 중앙 검은 창."""
    x1, y1, x2, y2 = center_window_coords(win_pct)
    print(f"\n  창 좌표: ({x1},{y1}) ~ ({x2},{y2})  창 크기: {win_pct:.1f}%")

    c(ser, _build(0x28, 0x60), "ALLCLR4: 전체 클리어")
    # BCOL4 [28H 64H]: 배경(base plane) 색 설정
    c(ser, _build(0x28, 0x64, 255, 255, 255, 8), "BCOL4(white): 배경 흰색")
    # 중앙 검은 창 등록
    c(ser, _build(0x28, 0x61, x1, y1, x2, y2), f"WINDOW4: 중앙 창 ({win_pct:.1f}%)")
    c(ser, _build(0x28, 0x62, 0, 0, 0, 8),       "WINCOL4(black): 중앙 창 검은색")
    c(ser, _build(0x24, 0x20, 0, 0), "EXPDN4(0,0): 렌더링")


# ── 방식 C: Full-field White → WINCLR4 → Black Center Window ─────────────────

def method_c_fullfield_then_window(ser: serial.Serial, win_pct: float) -> None:
    """full_field(white)로 전체 흰색 설정 후, window 플레인만 초기화하고 검은 창 추가."""
    x1, y1, x2, y2 = center_window_coords(win_pct)
    print(f"\n  창 좌표: ({x1},{y1}) ~ ({x2},{y2})  창 크기: {win_pct:.1f}%")

    # Step 1: 전체 화면 흰색 (full_field white)
    c(ser, _build(0x28, 0x60), "ALLCLR4: 전체 클리어")
    c(ser, _build(0x28, 0x61, 0, 0, _SCREEN_W-1, _SCREEN_H-1), "WINDOW4: 전체 화면 창 등록")
    c(ser, _build(0x28, 0x62, 255, 255, 255, 8), "WINCOL4(white): 전체 흰색")
    c(ser, _build(0x24, 0x20, 0, 0), "EXPDN4(0,0): 전체 흰색 렌더링", sleep=0.3)

    # Step 2: window 플레인만 클리어 (WINCLR4) → 베이스 패턴은 유지
    c(ser, _build(0x28, 0x63), "WINCLR4: 창 플레인만 클리어 (베이스 유지)")
    time.sleep(0.2)

    # Step 3: 중앙 검은 창 등록
    c(ser, _build(0x28, 0x61, x1, y1, x2, y2), f"WINDOW4: 중앙 창 ({win_pct:.1f}%)")
    c(ser, _build(0x28, 0x62, 0, 0, 0, 8),       "WINCOL4(black): 중앙 창 검은색")
    c(ser, _build(0x24, 0x20, 0, 0), "EXPDN4(0,0): 최종 렌더링")


# ── 방식 D: SPT4 fg/bg 분리 설정 후 EXPDN4 (패턴 프로그램 방식) ──────────────

def method_d_spt4_fg_bg(ser: serial.Serial, win_pct: float) -> None:
    """SPT4로 fg(black window)와 bg(white) 를 별도 설정.
    센터 정렬 패턴에서 쓴 방식의 응용 — fg를 검은 창, bg를 흰색으로.
    """
    x1, y1, x2, y2 = center_window_coords(win_pct)
    print(f"\n  창 좌표: ({x1},{y1}) ~ ({x2},{y2})  창 크기: {win_pct:.1f}%")

    # SPT4 [20H 2CH]: block 1=fg, block 18=bg
    c(ser, _build(0x20, 0x2C, 9999, 1,  0,   0,   0,   8), "SPT4 fg=black (block 1)")
    c(ser, _build(0x20, 0x2C, 9999, 18, 255, 255, 255, 8), "SPT4 bg=white (block 18)")
    # WINDOW4 창 등록 (fg 영역)
    c(ser, _build(0x28, 0x61, x1, y1, x2, y2), f"WINDOW4: 중앙 창")
    c(ser, _build(0x28, 0x62, 0, 0, 0, 8), "WINCOL4(black): 창 색상")
    c(ser, _build(0x24, 0x20, 9999, 0), "EXPDN4(9999,0): 프로그램 실행")


# ── 메인 ──────────────────────────────────────────────────────────────────────

def main() -> None:
    port = sys.argv[1] if len(sys.argv) > 1 else "/dev/ttyUSB0"
    print(f"포트: {port}  |  테스트할 창 크기 = 50% APL 기준 (흑창 side={70.7:.1f}%)")

    ser = serial.Serial(
        port, baudrate=38400, bytesize=serial.EIGHTBITS,
        parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
        rtscts=False, timeout=1.0,
    )
    time.sleep(0.3)

    if not enter_terminal(ser):
        ser.close()
        sys.exit(1)

    load_timing(ser)

    # ── Step 0: White Raster (Full-field White) 확인 ──────────────────────────
    print("\n" + "="*65)
    print(" Step 0: White Raster (전체 흰색) 출력")
    print("="*65)
    c(ser, _build(0x28, 0x60),                                     "ALLCLR4: 전체 클리어")
    c(ser, _build(0x28, 0x61, 0, 0, _SCREEN_W-1, _SCREEN_H-1),    "WINDOW4: 전체 화면 창 등록")
    c(ser, _build(0x28, 0x62, 255, 255, 255, 8),                   "WINCOL4(white): 흰색")
    c(ser, _build(0x24, 0x20, 0, 0),                               "EXPDN4(0,0): 렌더링")
    time.sleep(0.5)

    ans = input("\n  화면이 전체 흰색(White Raster)인가? [y=확인 / n=아님]: ").strip().lower()
    if ans != "y":
        print("  White Raster 출력 실패 — 장비/케이블 상태 확인 후 재시도하세요.")
        ser.close()
        sys.exit(1)
    print("  ✓ White Raster 확인 완료. 이제 Black Window 방식을 테스트합니다.")

    # 테스트할 창 크기 (APL 50% → 흑창 side = sqrt(0.5)*100 = 70.7%)
    WIN_PCT = 70.7

    methods = {
        "A": ("4-스트립 White Border + Black Center (현재 구현)", method_a_four_strips),
        "B": ("BCOL4 배경=White + WINDOW4 Black Center",          method_b_bcol4),
        "C": ("Full-field White → WINCLR4 → Black Window",       method_c_fullfield_then_window),
        "D": ("SPT4 fg=Black / bg=White + WINDOW4",              method_d_spt4_fg_bg),
    }

    print("\n" + "="*65)
    print(" 명암비 패턴 방식 테스트")
    print("="*65)
    print(" 각 방식을 실행한 뒤 화면을 확인하고 결과를 입력하세요.")
    print(" 기대 패턴: 흰 배경 전체 + 중앙 정사각형 검은 창\n")

    results: dict[str, str] = {}

    for key, (desc, fn) in methods.items():
        print(f"\n{'─'*65}")
        print(f" 방식 {key}: {desc}")
        print(f"{'─'*65}")
        fn(ser, WIN_PCT)
        time.sleep(0.5)

        ans = input(
            f"\n  화면 확인 → 올바른 패턴(흰 배경+검은 박스)인가?\n"
            f"  [y=맞음 / n=틀림 / p=부분적으로 맞음 / s=건너뜀]: "
        ).strip().lower()
        results[key] = ans

        if ans in ("y", "p"):
            print(f"  ★ 방식 {key} 동작 확인!")
            follow = input("  나머지 방식도 계속 테스트할까요? [y/n]: ").strip().lower()
            if follow != "y":
                break

        # 다음 테스트 전 초기화
        if key != list(methods.keys())[-1]:
            print("\n  다음 방식 전 화면 초기화 중...")
            c(ser, _build(0x28, 0x60), "ALLCLR4")
            c(ser, _build(0x28, 0x61, 0, 0, _SCREEN_W-1, _SCREEN_H-1), "WINDOW4(full)")
            c(ser, _build(0x28, 0x62, 128, 128, 128, 8), "WINCOL4(gray): 중간 회색으로 초기화")
            c(ser, _build(0x24, 0x20, 0, 0), "EXPDN4")
            time.sleep(1.0)

    ser.close()

    print("\n" + "="*65)
    print(" 테스트 결과 요약")
    print("="*65)
    for key, ans in results.items():
        desc = methods[key][0]
        mark = "★ 성공" if ans == "y" else ("△ 부분" if ans == "p" else ("✗ 실패" if ans == "n" else "건너뜀"))
        print(f"  방식 {key}: {mark}  —  {desc}")

    passed = [k for k, v in results.items() if v in ("y", "p")]
    if passed:
        print(f"\n  → 동작하는 방식: {', '.join(passed)}")
        print("  → 해당 방식으로 vg_generator.py의 show_window_patch를 수정하면 됩니다.")
    else:
        print("\n  → 동작하는 방식이 없습니다. 추가 조사 필요.")


if __name__ == "__main__":
    main()
