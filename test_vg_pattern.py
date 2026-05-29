"""VG-879 SPTS4 [20H 2AH] 패턴 선택 테스트.

SPTS4 커맨드 (매뉴얼 sec 2.11):
  STX(02H) + FDH + 20H + 2AH
  + 프로그램번호 + ,(2CH) + 패턴코드#1 + ,(2CH) + 패턴코드#2 ... + ETX(03H)

프로그램 번호 9999 = command work RAM (임시 즉시 실행용).
EXPDN4(9999, 0) 으로 실행.

패턴 코드 (Fig. 2-11-1):
  0=R, 1=G, 2=B, 3=INV, 8=Checker, 9=Aspect,
  10=Raster, 11=Moonscape, 12=Sweep,
  13=Ramp, 14=GrayScale, 15=ColorBar,
  19=Window, 25=Circle, 29=DOTS, 30=CROSS
"""

import serial
import time

# ---------------------------------------------------------------------------
# 포트 설정
# ---------------------------------------------------------------------------
PORT     = '/dev/ttyUSB0'
BAUDRATE = 38400

s = serial.Serial(PORT, BAUDRATE, timeout=1.0)
s.dtr = True
s.rts = True
time.sleep(0.5)

# ---------------------------------------------------------------------------
# 프로토콜 상수
# ---------------------------------------------------------------------------
STX = b'\x02'
ETX = b'\x03'
FD  = b'\xfd'
ENQ = b'\x05'
ACK = b'\x06'

# EXPDN4 프로그램 번호
PROG_4K60        = 2286   # 4K 60p 타이밍 (컬러바 프리셋)
PROG_WORK_RAM    = 9999   # command work RAM — SPTS4로 등록 후 즉시 실행


# ---------------------------------------------------------------------------
# 헬퍼
# ---------------------------------------------------------------------------
def _build(cmd1: int, cmd2: int, *params) -> bytes:
    """VG4 바이너리 프레임 빌더."""
    sep  = b'\x2C'
    body = sep.join(str(p).encode() for p in params)
    return STX + FD + bytes([cmd1, cmd2]) + body + ETX


def send(frame: bytes, label: str = "") -> bytes:
    s.reset_input_buffer()
    s.write(frame)
    time.sleep(0.5)
    resp = s.read(s.in_waiting or 64)
    if not resp:
        print(f"  [{label}] → (무응답)")
        return b''
    h = resp.hex()
    if b'\x11' in resp:
        idx  = resp.index(0x11)
        code = resp[idx + 1:].rstrip(b'\x03').decode('ascii', errors='replace')
        tag  = f"ESTS err={code}"
    elif b'\x06' in resp:
        tag = "ACK ✓"
    else:
        tag = "?"
    print(f"  [{label}] → [{tag}] {h}")
    return resp


def enter_terminal_mode():
    """ENQ → ACK 핸드셰이크로 터미널 모드 진입."""
    for _ in range(3):
        s.reset_input_buffer()
        s.write(ENQ)
        time.sleep(0.3)
        if s.in_waiting and s.read(1) == ACK:
            print("  터미널 모드 진입 성공")
            return True
    print("  터미널 모드 진입 실패 (계속 시도)")
    return False


# ---------------------------------------------------------------------------
# SPTS4 + EXPDN4 패턴 출력
# ---------------------------------------------------------------------------
def show_pattern(*codes: int, label: str = "") -> None:
    """SPTS4(9999, code...) → EXPDN4(9999, 0) 으로 즉시 패턴 출력.

    여러 코드를 전달하면 순서대로 레이아웃에 배치된다.
    단일 코드 권장 (멀티코드는 장비 프로그램 편집 의존적).
    """
    frame_spts4 = _build(0x20, 0x2A, PROG_WORK_RAM, *codes)
    frame_expdn = _build(0x24, 0x20, PROG_WORK_RAM, 0)
    send(frame_spts4, f"SPTS4(9999, {codes}) {label}")
    time.sleep(0.2)
    send(frame_expdn,  f"EXPDN4(9999,0) {label}")


# ---------------------------------------------------------------------------
# 4K 60p 타이밍 로드 (공통 초기화)
# ---------------------------------------------------------------------------
def load_timing():
    """EXPDN4(2286, 0) 으로 4K 60p 타이밍 로드."""
    print("\n  4K60p 타이밍 로드 중...")
    send(_build(0x24, 0x20, PROG_4K60, 0), "EXPDN4(2286,0)")
    time.sleep(1.5)


# ===========================================================================
# 메인 테스트
# ===========================================================================
print("=" * 60)
print("VG-879 SPTS4 패턴 테스트")
print("=" * 60)

enter_terminal_mode()
load_timing()

# ---------------------------------------------------------------------------
# 패턴 코드 테이블 (매뉴얼 Fig. 2-11-1 기반)
# ---------------------------------------------------------------------------
PATTERNS = [
    (15, "ColorBar    (현재 기본 초기화 패턴)"),
    (13, "Ramp        (계조 램프 — 대체 초기화 후보)"),
    (14, "GrayScale   (회색 계조)"),
    (10, "Raster      (단색 래스터)"),
    ( 8, "Checker     (체커보드)"),
    (12, "Sweep       (스윕)"),
    (11, "Moonscape   (문스케이프)"),
    (25, "Circle      (원형)"),
    (30, "CROSS       (크로스헤어)"),
    (29, "DOTS        (도트)"),
    ( 9, "Aspect      (종횡비 마커)"),
]

for code, desc in PATTERNS:
    print(f"\n{'─'*50}")
    print(f"  코드 {code:2d}: {desc}")
    show_pattern(code, label=desc.split()[0])
    time.sleep(1.5)
    ans = input(f"  화면 확인? (Enter=다음 / s=건너뜀 / q=종료) ").strip().lower()
    if ans == 'q':
        break

# ---------------------------------------------------------------------------
# 결과 확인: Ramp 이 컬러바보다 나은지 비교
# ---------------------------------------------------------------------------
print("\n" + "=" * 60)
print("비교 테스트: ColorBar vs Ramp (초기화 대체 후보)")
print("=" * 60)

input("\n[1] ColorBar (Enter 누르면 출력)... ")
load_timing()  # EXPDN4(2286,0) → 기본 컬러바
show_pattern(15, label="ColorBar")
input("    ColorBar 확인 완료? (Enter) ")

input("\n[2] Ramp (Enter 누르면 출력)... ")
load_timing()
show_pattern(13, label="Ramp")
input("    Ramp 확인 완료? (Enter) ")

# ---------------------------------------------------------------------------
# vg_generator.py 적용 방법 안내
# ---------------------------------------------------------------------------
print("""
=== vg_generator.py 반영 방법 ===

Ramp 가 정상 출력됐다면, 초기화 시퀀스를 아래처럼 바꾸면 된다:

  기존: EXPDN4(2286, 0)           # 컬러바 프로그램 로드
  변경: EXPDN4(2286, 1)           # 타이밍만 로드 (패턴 변경 없음)
        SPTS4(9999, 13)           # Ramp 패턴을 work RAM에 등록
        EXPDN4(9999, 0)           # Ramp 즉시 실행

  또는 간단히 vg_generator.py 상단의:
    _PROG_INIT = _PROG_4K60_RAMP  ← 이 방식은 별도 램프 프로그램 번호 필요

  SPTS4 방식이 프로그램 번호 불필요하므로 더 권장.
""")

s.close()
print("완료.")
