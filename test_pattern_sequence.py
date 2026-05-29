"""패턴 연속 출력 검증 스크립트.

테스트 1: APL 크기 변경 확인 (lum_loading 재현)
  - WINCLR4 방식: 매 스텝 라인 메모리 클리어 후 새 크기 등록
  - 예상: 각 스텝에서 패턴 크기가 실제로 바뀜

테스트 2: 색상 연속 전환 확인 (gamut 재현)
  - _timing_loaded 방식: 타이밍 1회 로드 후 패턴만 교체
  - 예상: R→G→B 전환 시 컬러바 깜빡임 없음

모든 커맨드 응답(ACK/ESTS) 출력.
"""

import serial
import time

PORT     = '/dev/ttyUSB0'
BAUDRATE = 38400
SCREEN_W = 3840
SCREEN_H = 2160

s = serial.Serial(PORT, BAUDRATE, timeout=1.0)
s.dtr = True
s.rts = True
time.sleep(0.5)

STX = b'\x02'
ETX = b'\x03'
FD  = b'\xfd'
SEP = b'\x2C'
ENQ = b'\x05'
ACK = b'\x06'


def build(cmd1, cmd2, *params):
    parts = SEP.join(str(p).encode() for p in params)
    return STX + FD + bytes([cmd1, cmd2]) + parts + ETX


def send(frame, label):
    s.reset_input_buffer()
    s.write(frame)
    time.sleep(0.5)
    resp = s.read(s.in_waiting or 64)
    if not resp:
        tag = "(무응답)"
    elif b'\x11' in resp:
        idx = resp.index(0x11)
        code = resp[idx+1:].rstrip(b'\x03').decode('ascii', errors='replace')
        tag = f"ESTS err={code}"
    elif b'\x06' in resp:
        tag = "ACK ✓"
    else:
        tag = f"? {resp.hex()}"
    print(f"  {label:52s} → {tag}")
    return tag


def enter_terminal():
    for _ in range(3):
        s.reset_input_buffer()
        s.write(ENQ)
        time.sleep(0.3)
        if s.in_waiting and s.read(1) == ACK:
            print("  터미널 모드 OK")
            return
    print("  터미널 모드 실패")


def load_colorbar():
    send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0) 컬러바 로드")
    time.sleep(2.0)


def allclr():
    send(build(0x28, 0x60), "ALLCLR4 [28H 60H]")
    time.sleep(0.3)


def show_window(w_pct, h_pct, r, g, b, bg_r=0, bg_g=0, bg_b=0):
    """ALLCLR4 → BCOL4(배경) → WINDOW4(중앙) → WINCOL4(색상) → EXPDN4(0,0)"""
    w = max(1, int(SCREEN_W * w_pct / 100))
    h = max(1, int(SCREEN_H * h_pct / 100))
    x1 = (SCREEN_W - w) // 2
    y1 = (SCREEN_H - h) // 2
    x2 = x1 + w - 1
    y2 = y1 + h - 1
    allclr()
    send(build(0x28, 0x71, bg_r, bg_g, bg_b, 8), f"  BCOL4({bg_r},{bg_g},{bg_b})")
    time.sleep(0.3)
    send(build(0x28, 0x61, x1, y1, x2, y2), f"  WINDOW4({x1},{y1},{x2},{y2}) {w}×{h}px")
    time.sleep(0.3)
    send(build(0x28, 0x62, r, g, b, 8), f"  WINCOL4({r},{g},{b})")
    time.sleep(0.3)
    send(build(0x24, 0x20, 0, 0), "  EXPDN4(0,0)")
    time.sleep(0.5)


def show_fullfield(r, g, b):
    """ALLCLR4 → WINDOW4(전체) → WINCOL4 → EXPDN4(0,0)  (타이밍 재로드 없음)"""
    allclr()
    send(build(0x28, 0x61, 0, 0, SCREEN_W-1, SCREEN_H-1), "  WINDOW4(0,0,3839,2159) 전체")
    time.sleep(0.3)
    send(build(0x28, 0x62, r, g, b, 8), f"  WINCOL4({r},{g},{b})")
    time.sleep(0.3)
    send(build(0x24, 0x20, 0, 0), "  EXPDN4(0,0)")
    time.sleep(0.5)


# ─────────────────────────────────────────────────────────────────────────────
print("=" * 65)
print("패턴 연속 출력 검증")
print("=" * 65)
enter_terminal()

# ─────────────────────────────────────────────────────────────────────────────
print("\n[테스트 1] APL 크기 변경 (lum_loading 재현)")
print("  방식: 타이밍 1회 로드 → 각 APL마다 WINCLR4 → WINDOW4(새 크기) → EXPDN4")
print("-" * 65)

load_colorbar()

apl_steps = [10, 25, 50, 100, 25, 10]
for apl in apl_steps:
    import math
    side = math.sqrt(apl / 100.0) * 100.0
    print(f"\n  APL {apl:3d}%  W={side:.1f}%  H={side:.1f}%")
    show_window(side, side, 255, 255, 255)
    result = input(f"    APL {apl}% 흰색 창 크기가 올바른가요? (y/n): ").strip().lower()
    if result != 'y':
        print(f"    ✗ APL {apl}% 실패")
        break
    print(f"    ✓ APL {apl}% OK")

# ─────────────────────────────────────────────────────────────────────────────
print("\n[테스트 2] 색상 연속 전환 (gamut 재현 — 타이밍 재로드 없음)")
print("  방식: 타이밍 1회 로드 → R/G/B/W 각각 WINCLR4 → WINDOW4(전체) → WINCOL4")
print("  확인 포인트: 색상 전환 시 컬러바 깜빡임이 없어야 함")
print("-" * 65)

load_colorbar()

colors = [
    ("빨강", 255,   0,   0),
    ("초록",   0, 255,   0),
    ("파랑",   0,   0, 255),
    ("흰색", 255, 255, 255),
]
for name, r, g, b in colors:
    print(f"\n  → {name} 전체 화면")
    show_fullfield(r, g, b)
    result = input(f"    {name} 전체 화면 OK + 깜빡임 없음? (y/n): ").strip().lower()
    if result != 'y':
        print(f"    ✗ {name} 실패")
        break
    print(f"    ✓ {name} OK")

# ─────────────────────────────────────────────────────────────────────────────
print("\n컬러바 복귀...")
load_colorbar()
s.close()
print("완료.")
