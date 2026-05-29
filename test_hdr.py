"""HDR 출력 진단 테스트 v2 — 이전 실패 원인 분석 후 수정된 케이스.

[이전 실패 원인 추정]
  Max Disp Mastering: 4000 nits → 현재 코드가 4000*10000=40,000,000 전송
  매뉴얼: "0"-"65535" (5 bytes 한계) → 8자리 숫자로 ESTS 에러 발생

[매뉴얼 정확한 단위]
  Max Disp Mastering : 0-65535, 단위 = 1 cd/m²  (4000 nits → 4000)
  Min Disp Mastering : 0-65535, 단위 = 0.0001 cd/m² (*주: 10,000배)  (0.0001 nit → 1)
  Content Light LV   : 0-65535, 단위 = 1 cd/m²
  Frame-ave Light LV : 0-65535, 단위 = 1 cd/m²
  Disp Primaries     : 0-100000, 단위 = 실제값×100,000 (BT.2020 기준)

[이번 케이스]
  H. 올바른 Max 값(4000) + SHDR4 단독
  I. 올바른 Max 값 + SHDMI4 + SHDR4
  J. 올바른 Max 값 + SHDMI4 + SHDR4, EXPDN4 없음 (설정만)
  K. 올바른 Max 값 + SHDR4(prog=9999) + EXPDN4(9999,0)
  L. 메타데이터 전부 0 (최소한) + SHDR4
  M. EXPDN4(2286,0) 없이 바로 SHDR4 → EXPDN4(0,0)
  N. SHDR4 ON 상태에서 EXPDN4(2286,0) 마지막으로
"""

import math
import serial
import time

PORT     = '/dev/ttyUSB0'
BAUDRATE = 38400

s = serial.Serial(PORT, BAUDRATE, timeout=1.0)
s.dtr = True
s.rts = True
time.sleep(0.5)

STX = b'\x02'
ETX = b'\x03'
ENQ = b'\x05'
ACK = b'\x06'


# ─── 프레임 빌더 / 송수신 ─────────────────────────────────────────────────────
def build(cmd1: int, cmd2: int, *params) -> bytes:
    body = b','.join(str(p).encode() for p in params)
    return STX + b'\xFD' + bytes([cmd1, cmd2]) + body + ETX


def send(frame: bytes, label: str = "") -> str:
    s.reset_input_buffer()
    s.write(frame)
    time.sleep(0.4)
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
    print(f"    {label:<55} → {tag}")
    return tag


def enter_terminal():
    for _ in range(3):
        s.reset_input_buffer()
        s.write(ENQ)
        time.sleep(0.3)
        if s.in_waiting and s.read(1) == ACK:
            print("  터미널 모드 OK\n")
            return True
    print("  터미널 모드 실패\n")
    return False


def load_timing(prog: int = 2286, sleep: float = 1.2):
    send(build(0x24, 0x20, prog, 0), f"EXPDN4({prog},0)")
    time.sleep(sleep)


def show_white_10pct():
    """흰색 10% APL 윈도우 출력 (10-bit)."""
    side = math.sqrt(0.10)
    w = max(1, int(3840 * side))
    h = max(1, int(2160 * side))
    x1 = (3840 - w) // 2
    y1 = (2160 - h) // 2
    x2 = x1 + w - 1
    y2 = y1 + h - 1
    send(build(0x28, 0x60),                       "ALLCLR4")
    time.sleep(0.2)
    send(build(0x28, 0x71, 0, 0, 0, 10),          "BCOL4(black 10bit)")
    time.sleep(0.2)
    send(build(0x28, 0x61, x1, y1, x2, y2),       f"WINDOW4 10%APL {x1},{y1}~{x2},{y2}")
    time.sleep(0.2)
    send(build(0x28, 0x62, 1023, 1023, 1023, 10), "WINCOL4(white 10bit)")
    time.sleep(0.2)
    send(build(0x24, 0x20, 0, 0),                 "EXPDN4(0,0) 렌더링")
    time.sleep(0.8)


def restore():
    """SDR 복귀."""
    print("  [SDR 복귀]")
    # SHDR4 Off
    send(build(0x20, 0xC5,
               0, 0,        # prog=0, Off
               7, 1, 0, 0,  # Type, Version, EOTF=SDR, MetaID=0
               70800, 29200, 17000, 79700, 13100, 4600,
               31270, 32900,
               4000, 1, 4000, 400,  # ← Max=4000 (올바른 값)
               0),
         "SHDR4 Off (SDR)")
    time.sleep(0.2)
    send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
    time.sleep(0.3)
    load_timing(2286, sleep=1.5)
    print()


# ─── SHDR4 헬퍼 (올바른 Max 값 사용) ─────────────────────────────────────────
# Max Disp Mastering: 단위 1 cd/m²  → 4000 nits = 4000
# Min Disp Mastering: 단위 0.0001 cd/m² (×10000) → 0.0001 nit = 1
def shdr4_frame(prog: int, on_off: int, eotf: int,
                max_lum: int = 4000, min_lum_x10000: int = 1,
                max_cll: int = 4000, max_fall: int = 400) -> bytes:
    return build(
        0x20, 0xC5,
        prog, on_off,
        7, 1,                                              # Type=7, Version=1
        eotf,
        0,                                                 # Metadata ID=0
        70800, 29200, 17000, 79700, 13100, 4600,           # BT.2020 primaries ×100000
        31270, 32900,                                      # D65 white point ×100000
        max_lum, min_lum_x10000, max_cll, max_fall,        # ← Max=4000 (수정됨)
        0,                                                 # Data Type=HDMI
    )


def shdmi4_ycbcr422_10bit():
    # YCbCr422, 10bit, BT.2020
    return build(0x20, 0x36, 0, 1, 1, 0, 2, 0, 0)


# ─── 진입 ─────────────────────────────────────────────────────────────────────
print("=" * 65)
print("HDR 출력 진단 테스트 v2  (Max 값 버그 수정)")
print("=" * 65)
enter_terminal()

# =============================================================================
# 케이스 H: SHDR4 단독, Max=4000(올바른 값), EOTF=ST2084
# =============================================================================
print("─" * 65)
print("[케이스 H] SHDR4 단독, Max=4000(수정), prog=0, EOTF=ST2084(2)")
print("─" * 65)
load_timing()
send(shdr4_frame(0, 1, 2), "SHDR4(On, ST2084, Max=4000)")
time.sleep(0.2)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
time.sleep(0.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스H → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 I: SHDMI4 + SHDR4(Max=4000)
# =============================================================================
print("─" * 65)
print("[케이스 I] SHDMI4(YCbCr422,10bit,BT.2020) + SHDR4(Max=4000)")
print("─" * 65)
load_timing()
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
send(shdr4_frame(0, 1, 2), "SHDR4(On, ST2084, Max=4000)")
time.sleep(0.2)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
time.sleep(0.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스I → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 J: SHDR4(prog=9999) + EXPDN4(9999,0)  — work RAM 방식
# =============================================================================
print("─" * 65)
print("[케이스 J] SHDR4 prog=9999 + EXPDN4(9999,0)  (work RAM)")
print("─" * 65)
load_timing()
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
send(shdr4_frame(9999, 1, 2), "SHDR4(prog=9999, On, ST2084)")
time.sleep(0.2)
send(build(0x24, 0x20, 9999, 0), "EXPDN4(9999,0)")
time.sleep(0.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스J → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 K: EXPDN4 없이 SHDR4만 (EXPDN4 없는 패턴)
# =============================================================================
print("─" * 65)
print("[케이스 K] SHDR4 설정 후 EXPDN4 호출 없이 패턴만 출력")
print("─" * 65)
load_timing()
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
send(shdr4_frame(0, 1, 2), "SHDR4(On, ST2084, Max=4000)")
time.sleep(0.5)
# EXPDN4(0,0) 없이 바로 패턴
side = math.sqrt(0.10)
w = max(1, int(3840 * side)); h = max(1, int(2160 * side))
x1 = (3840 - w) // 2; y1 = (2160 - h) // 2
send(build(0x28, 0x60), "ALLCLR4")
time.sleep(0.2)
send(build(0x28, 0x71, 0, 0, 0, 10), "BCOL4(black)")
time.sleep(0.2)
send(build(0x28, 0x61, x1, y1, x1+w-1, y1+h-1), "WINDOW4 10%APL")
time.sleep(0.2)
send(build(0x28, 0x62, 1023, 1023, 1023, 10), "WINCOL4(white)")
time.sleep(0.2)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0) 렌더링")
time.sleep(0.8)
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스K → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 L: 메타데이터 최소화 (Primaries/WhitePoint 전부 0)
# =============================================================================
print("─" * 65)
print("[케이스 L] 메타데이터 최소화 (Primaries=0, EOTF=ST2084)")
print("─" * 65)
load_timing()
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
send(build(0x20, 0xC5,
           0, 1, 7, 1, 2, 0,   # prog=0, On, Type=7, Ver=1, EOTF=ST2084, MetaID=0
           0, 0, 0, 0, 0, 0,   # Primaries 전부 0
           0, 0,                # White point 전부 0
           1000, 1, 1000, 400,  # Max=1000(무난한 값), Min=1, CLL=1000, FALL=400
           0),
     "SHDR4(On, ST2084, Primaries=0, Max=1000)")
time.sleep(0.2)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
time.sleep(0.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스L → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 M: EXPDN4(2286,0) 없이 시작 → SHDR4 → EXPDN4(0,0)
#            (타이밍 로드가 HDR 설정을 초기화하는지 확인)
# =============================================================================
print("─" * 65)
print("[케이스 M] 타이밍 로드(EXPDN4 2286) 없이 바로 SHDR4 설정")
print("─" * 65)
# 타이밍 로드 하지 않고 바로
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
send(shdr4_frame(0, 1, 2), "SHDR4(On, ST2084, Max=4000) — 타이밍 로드 없음")
time.sleep(0.2)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
time.sleep(0.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스M → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 N: SHDR4 → EXPDN4(2286,0) 마지막에 (순서 반전)
# =============================================================================
print("─" * 65)
print("[케이스 N] SHDR4 설정 후 EXPDN4(2286,0)으로 타이밍 로드 마지막에")
print("─" * 65)
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
send(shdr4_frame(0, 1, 2), "SHDR4(On, ST2084, Max=4000)")
time.sleep(0.2)
# 타이밍 로드를 마지막에
send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0) ← 마지막에 타이밍")
time.sleep(1.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스N → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 O: On/Off 파라미터 생략 — 혹시 이 펌웨어에 On/Off 필드가 없는지 확인
#   구조: ProgNO, Type, Version, EOTF, MetaID, primaries..., Max, Min, CLL, FALL, DataType
#   (On/Off 없는 버전으로 테스트)
# =============================================================================
print("─" * 65)
print("[케이스 O] On/Off 파라미터 없이 전송 (펌웨어 버전 차이 확인)")
print("─" * 65)
load_timing()
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
# On/Off 없이: ProgNO, Type, Version, EOTF, MetaID, ...
send(build(0x20, 0xC5,
           0,            # Program NO
           7, 1,         # Type=7, Version=1  (On/Off 없음)
           2,            # EOTF=ST2084
           0,            # Metadata ID
           70800, 29200, 17000, 79700, 13100, 4600,
           31270, 32900,
           4000, 1, 4000, 400,
           0),
     "SHDR4 On/Off 파라미터 없이 (Type 바로 시작)")
time.sleep(0.2)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
time.sleep(0.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스O → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
# 케이스 P: SHDR4 On/Off=1 + Type/Version 없이 (다른 펌웨어 구조)
#   구조: ProgNO, On/Off, EOTF, MetaID, primaries...
#   (Type, Version 없는 버전)
# =============================================================================
print("─" * 65)
print("[케이스 P] Type/Version 생략 버전 (ProgNO, On/Off, EOTF, MetaID, ...)")
print("─" * 65)
load_timing()
send(shdmi4_ycbcr422_10bit(), "SHDMI4(YCbCr422, 10bit, BT.2020)")
time.sleep(0.3)
send(build(0x20, 0xC5,
           0,            # Program NO
           1,            # On/Off=ON
           2,            # EOTF=ST2084 (Type/Version 없음)
           0,            # Metadata ID
           70800, 29200, 17000, 79700, 13100, 4600,
           31270, 32900,
           4000, 1, 4000, 400,
           0),
     "SHDR4 Type/Version 없이 (ProgNO,On/Off,EOTF,...)")
time.sleep(0.2)
send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
time.sleep(0.5)
show_white_10pct()
ans = input("  HDR 아이콘? (y/n/s): ").strip().lower()
print(f"  케이스P → {'성공' if ans=='y' else '실패' if ans=='n' else '스킵'}\n")
restore()

# =============================================================================
print("=" * 65)
print("완료. 컬러바 복귀...")
load_timing(2286, sleep=1.5)
s.close()
print("종료.")
