"""show_full_field 시퀀스 단계별 진단 스크립트.

각 VG 커맨드의 응답을 출력하여 어느 단계에서 문제가 발생하는지 확인한다.
현재 production driver (vg_generator.py) 의 show_full_field() 와
정확히 동일한 순서로 커맨드를 전송하되, 모든 응답을 콘솔에 출력한다.

접근법 A : 현재 production 방식 (WINCLR4 + WINDOW4 full screen + WINCOL4 + EXP0)
접근법 B : test_vg.py 방식   (EXP2286 → tiny black window → EXP0 → WINDOW4 → WINCOL4 → EXP0)
두 가지를 모두 실행하여 화면 출력 결과를 비교한다.
"""

import serial
import time

# ---------------------------------------------------------------------------
PORT     = '/dev/ttyUSB0'
BAUDRATE = 38400
# ---------------------------------------------------------------------------

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


def build(cmd1: int, cmd2: int, *params) -> bytes:
    parts = SEP.join(str(p).encode() for p in params)
    return STX + FD + bytes([cmd1, cmd2]) + parts + ETX


def send(frame: bytes, label: str) -> str:
    s.reset_input_buffer()
    s.write(frame)
    time.sleep(0.5)
    resp = s.read(s.in_waiting or 64)
    if not resp:
        tag = "(무응답)"
    elif b'\x11' in resp:
        idx  = resp.index(0x11)
        code = resp[idx + 1:].rstrip(b'\x03').decode('ascii', errors='replace')
        tag  = f"ESTS err={code}"
    elif b'\x06' in resp:
        tag = "ACK ✓"
    else:
        tag = f"? {resp.hex()}"
    print(f"  {label:50s} → {tag}")
    return tag


def enter_terminal_mode():
    for _ in range(3):
        s.reset_input_buffer()
        s.write(ENQ)
        time.sleep(0.3)
        if s.in_waiting and s.read(1) == ACK:
            print("  터미널 모드 진입 OK")
            return
    print("  터미널 모드 실패 — 계속 진행합니다")


# ===========================================================================
print("=" * 65)
print("show_full_field 진단 스크립트")
print("=" * 65)
enter_terminal_mode()

# ---------------------------------------------------------------------------
# 접근법 A — 현재 production driver 방식
# ---------------------------------------------------------------------------
print("\n[A] Production 방식: EXPDN4(2286,0) → WINCLR4 → WINDOW4(full) → WINCOL4(red) → EXPDN4(0,0)")
print("-" * 65)

send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0)  타이밍 로드")
time.sleep(2.0)

send(build(0x28, 0x63), "WINCLR4  [28H 63H]  윈도우 메모리 클리어")
time.sleep(0.3)

send(build(0x28, 0x61, 0, 0, 3839, 2159), "WINDOW4(0,0,3839,2159)  풀스크린")
time.sleep(0.3)

send(build(0x28, 0x62, 255, 0, 0, 8), "WINCOL4(255,0,0,8)  빨강")
time.sleep(0.3)
send(build(0x28, 0x62, 0, 255, 0, 8), "WINCOL4(0,255,0,8)  green")
time.sleep(0.3)
send(build(0x28, 0x62, 0, 0, 255, 8), "WINCOL4(0,0,255,8)  blue")
time.sleep(0.3) 
send(build(0x28, 0x62, 255, 255, 255, 8), "WINCOL4(255,255,255,8)  white")
time.sleep(0.3)
send(build(0x28, 0x62, 0, 0, 0, 8), "WINCOL4(0,0,0,8)  black")

send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)  실행")
time.sleep(1.0)

r_a = input("\n  [A] 화면이 빨간 전체화면으로 바뀌었나요? (y/n): ").strip().lower()

# ---------------------------------------------------------------------------
# 컬러바로 복귀
# ---------------------------------------------------------------------------
print("\n  컬러바로 복귀...")
send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0)  컬러바 복귀")
time.sleep(2.0)

# ---------------------------------------------------------------------------
# 접근법 B — test_vg.py 방식 (WINCLR4 없음, 작은 검정 창으로 초기화)
# ---------------------------------------------------------------------------
print("\n[B] test_vg.py 방식: EXP2286 → WINDOW4(0,0,0,0,black) → EXP0 → WINDOW4(full) → WINCOL4(red) → EXP0")
print("-" * 65)

send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0)  타이밍 로드")
time.sleep(2.0)

# 작은 더미 창으로 윈도우 메모리 초기화 (test_vg.py reset_windraw 방식)
send(build(0x28, 0x61, 0, 0, 0, 0), "WINDOW4(0,0,0,0)  더미 1×1 흑색 창")
time.sleep(0.3)
send(build(0x28, 0x62, 0, 0, 0, 8), "WINCOL4(0,0,0,8)  흑색")
time.sleep(0.3)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)  더미 창 적용")
time.sleep(0.5)

# 실제 패턴 전송
send(build(0x28, 0x61, 0, 0, 3839, 2159), "WINDOW4(0,0,3839,2159)  풀스크린")
time.sleep(0.3)
send(build(0x28, 0x62, 255, 0, 0, 8), "WINCOL4(255,0,0,8)  빨강")
time.sleep(0.3)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)  실행")
time.sleep(1.0)

r_b = input("\n  [B] 화면이 빨간 전체화면으로 바뀌었나요? (y/n): ").strip().lower()

# ---------------------------------------------------------------------------
# 접근법 C — WINCOL4 파라미터 없이 3개만 전송 (bit_mode 제외)
# ---------------------------------------------------------------------------
print("\n  컬러바로 복귀...")
send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0)  컬러바 복귀")
time.sleep(2.0)

print("\n[C] WINCOL4 3파라미터 (bit_mode 제외): WINCLR4 → WINDOW4(full) → WINCOL4(255,0,0) → EXPDN4(0,0)")
print("-" * 65)

send(build(0x28, 0x63), "WINCLR4")
time.sleep(0.3)
send(build(0x28, 0x61, 0, 0, 3839, 2159), "WINDOW4(0,0,3839,2159)  풀스크린")
time.sleep(0.3)
send(build(0x28, 0x62, 255, 0, 0), "WINCOL4(255,0,0)  빨강 (bit_mode 없음)")
time.sleep(0.3)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)  실행")
time.sleep(1.0)

r_c = input("\n  [C] 화면이 빨간 전체화면으로 바뀌었나요? (y/n): ").strip().lower()

# ---------------------------------------------------------------------------
# 결과 요약
# ---------------------------------------------------------------------------
print("\n" + "=" * 65)
print("진단 결과 요약")
print("=" * 65)
print(f"  [A] Production 방식 (WINCLR4 사용)        : {'성공 ✓' if r_a == 'y' else '실패 ✗'}")
print(f"  [B] test_vg.py 방식 (더미 창 초기화)       : {'성공 ✓' if r_b == 'y' else '실패 ✗'}")
print(f"  [C] WINCOL4 bit_mode 파라미터 제외         : {'성공 ✓' if r_c == 'y' else '실패 ✗'}")

if r_a != 'y' and r_b == 'y':
    print("\n→ WINCLR4 [28H 63H] 가 이 장비에서 동작하지 않는 것으로 보입니다.")
    print("  test_vg.py 방식(더미 창 초기화)으로 production 코드를 수정해야 합니다.")
elif r_a != 'y' and r_b != 'y' and r_c == 'y':
    print("\n→ WINCOL4 bit_mode 파라미터(4번째 인자)가 문제입니다.")
    print("  WINCOL4 를 r,g,b 3개 파라미터만 전송하도록 수정해야 합니다.")
elif r_a == 'y':
    print("\n→ Production 방식이 정상 동작합니다.")
    print("  show_full_field 자체는 문제없습니다. 다른 원인을 확인하세요.")
else:
    print("\n→ 세 방식 모두 실패. EXPDN4(0,0) 또는 WINDOW4 좌표가 문제일 수 있습니다.")

# ---------------------------------------------------------------------------
print("\n컬러바로 최종 복귀...")
send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0)")
s.close()
print("완료.")
