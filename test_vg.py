"""VG-879 manual serial test — WINDOW4 [28H 61H] APL window verification
+ HDR10 (ST2084) output toggle test.

WINDOW4 frame format (from manual section 3.32):
  STX(02H) + FDH + 28H + 61H
  + X1(0-4095) + ,(2CH) + Y1(0-4095) + ,(2CH)
  + X2(0-4095) + ,(2CH) + Y2(0-4095)
  + ETX(03H)

Screen: 3840 × 2160 (4K UHD)
  Full screen: X1=0, Y1=0, X2=3839, Y2=2159  (all within 0-4095 range)

APL window formula (square, centred):
  side_px_W = int(3840 * sqrt(apl/100))
  side_px_H = int(2160 * sqrt(apl/100))
  X1 = (3840 - side_px_W) // 2
  Y1 = (2160 - side_px_H) // 2
  X2 = X1 + side_px_W - 1
  Y2 = Y1 + side_px_H - 1

HDR 커맨드 순서 (section 2.140 SHDR4):
  1. EXPDN4(2286, 0)       : 4K60p 타이밍 로드
  2. SHDMI4 [20H 36H]      : HDMI 출력 포맷 (YCbCr422, 10-bit, BT.2020)
  3. SIF4   [20H 38H]      : AVI InfoFrame (BT.2020 colorimetry)
  4. SHDR4  [20H C5H]      : Dynamic Range/Mastering InfoFrame
       On/Off=1, EOTF=2 (ST2084) → HDR ON
       On/Off=0, EOTF=0 (SDR)   → HDR OFF
  5. EXPDN4(0, 0)          : 설정 반영
"""

import math
import serial
import time

# ---------------------------------------------------------------------------
# Serial port setup
# ---------------------------------------------------------------------------
s = serial.Serial('/dev/ttyUSB0', 38400, timeout=1.0)
s.dtr = True
s.rts = True
time.sleep(0.5)

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
STX = b'\x02'
ETX = b'\x03'
FD  = b'\xfd'

EXP2286 = STX + FD + b'\x24\x20' + b'2286,0' + ETX
EXP0    = STX + FD + b'\x24\x20' + b'0,0'    + ETX


def send(frame: bytes, label: str = "") -> bytes:
    s.reset_input_buffer()
    s.write(frame)
    time.sleep(0.5)
    resp = s.read(s.in_waiting or 64)
    if not resp:
        print(f"  {label}: (none)")
        return b''
    h = resp.hex()
    if b'\x11' in resp:
        idx = resp.index(0x11)
        code = resp[idx + 1:].rstrip(b'\x03').decode('ascii', errors='replace')
        tag = f"ESTS err={code}"
    elif b'\x06' in resp:
        tag = "ACK"
    else:
        tag = "?"
    print(f"  {label}: [{tag}] {h}")
    return resp


def window4(x1: int, y1: int, x2: int, y2: int) -> bytes:
    """Build WINDOW4 [FD 28H 61H] frame per manual section 3.32.
    Coordinates must be in range 0-4095.
    """
    assert 0 <= x1 <= x2 <= 4095, f"WINDOW4 X out of range: {x1},{x2}"
    assert 0 <= y1 <= y2 <= 4095, f"WINDOW4 Y out of range: {y1},{y2}"
    payload = f"{x1},{y1},{x2},{y2}".encode()
    return STX + FD + b'\x28\x61' + payload + ETX


def wincol4(r: int, g: int, b: int, bits: int = 8) -> bytes:
    """Build WINCOL4 [FD 28H 62H] frame."""
    payload = f"{r},{g},{b},{bits}".encode()
    return STX + FD + b'\x28\x62' + payload + ETX


def apl_window(apl_pct: float) -> tuple[int, int, int, int]:
    """Return (x1,y1,x2,y2) for a centred square window at the given APL %.

    side_W = int(3840 * sqrt(apl/100))
    side_H = int(2160 * sqrt(apl/100))
    Coordinates are inclusive on both endpoints per WINDOW4 spec.
    """
    side = math.sqrt(apl_pct / 100.0)
    w = int(3840 * side)
    h = int(2160 * side)
    w = max(1, w)
    h = max(1, h)
    x1 = (3840 - w) // 2
    y1 = (2160 - h) // 2
    x2 = x1 + w - 1
    y2 = y1 + h - 1
    return x1, y1, x2, y2


# ---------------------------------------------------------------------------
# Reset helper: reload timing, overwrite lingering WINDRAW with 1×1 black
# ---------------------------------------------------------------------------
def reset_windraw():
    send(EXP2286, "EXP2286 reload")
    time.sleep(1.5)
    send(window4(0, 0, 0, 0), "WINDOW4 clear (1×1)")
    send(wincol4(0, 0, 0),    "WINCOL4 black")
    send(EXP0,                 "EXP0 flush")
    time.sleep(0.5)


# ---------------------------------------------------------------------------
# Test 1 – centre alignment box 1000×1000
# ---------------------------------------------------------------------------
print("\n=== Test 1: 센터 정렬 박스 1000×1000 ===")
W_BOX, H_BOX = 1000, 1000
X1 = (3840 - W_BOX) // 2          # 1420
Y1 = (2160 - H_BOX) // 2          # 580
X2 = X1 + W_BOX - 1               # 2419  (inclusive)
Y2 = Y1 + H_BOX - 1               # 1579  (inclusive)
print(f"  WINDOW4 coords: ({X1},{Y1}) ~ ({X2},{Y2})  →  {W_BOX}×{H_BOX} px")

reset_windraw()
send(window4(X1, Y1, X2, Y2), f"WINDOW4 ({X1},{Y1},{X2},{Y2})")
send(wincol4(255, 255, 255),   "WINCOL4 white")
send(EXP0,                     "EXP0")
time.sleep(2.0)
input("센터 흰 박스 보임? (Enter) ")

# ---------------------------------------------------------------------------
# Test 2 – 10 % APL window (lum-swing pattern)
# ---------------------------------------------------------------------------
print("\n=== Test 2: 10 % APL 윈도우 ===")
x1, y1, x2, y2 = apl_window(10.0)
w_px = x2 - x1 + 1
h_px = y2 - y1 + 1
actual_apl = (w_px * h_px) / (3840 * 2160) * 100
print(f"  WINDOW4 coords: ({x1},{y1}) ~ ({x2},{y2})")
print(f"  Size: {w_px}×{h_px} px  |  APL = {actual_apl:.2f}%")

reset_windraw()
send(window4(x1, y1, x2, y2), f"WINDOW4 10%APL ({x1},{y1},{x2},{y2})")
send(wincol4(255, 255, 255),   "WINCOL4 white")
send(EXP0,                     "EXP0")
time.sleep(2.0)
input("10% APL 흰 박스 보임? (Enter) ")

# ---------------------------------------------------------------------------
# Test 3 – APL sweep (optional)
# ---------------------------------------------------------------------------
run_sweep = input("\nAPL sweep 테스트 실행? (y/N) ").strip().lower() == 'y'
if run_sweep:
    for apl in [100, 50, 25, 10, 5, 1]:
        x1, y1, x2, y2 = apl_window(float(apl))
        w_px = x2 - x1 + 1
        h_px = y2 - y1 + 1
        actual = (w_px * h_px) / (3840 * 2160) * 100
        print(f"\n  APL={apl}%  WINDOW4({x1},{y1},{x2},{y2})  {w_px}×{h_px}  actual={actual:.2f}%")
        reset_windraw()
        send(window4(x1, y1, x2, y2), f"WINDOW4 {apl}%APL")
        send(wincol4(255, 255, 255),   "WINCOL4 white")
        send(EXP0,                     "EXP0")
        time.sleep(2.0)
        input(f"  APL {apl}% 박스 보임? (Enter) ")

# ---------------------------------------------------------------------------
# Test 4 – HDR10 (ST2084) 출력 전환
# ---------------------------------------------------------------------------
def build(*args) -> bytes:
    """STX + FD + CMD1 + CMD2 + ASCII params(comma-sep) + ETX."""
    cmd1, cmd2, *params = args
    parts = b','.join(str(p).encode() for p in params)
    return STX + FD + bytes([cmd1, cmd2]) + parts + ETX


def set_hdr10(on: bool, max_lum: int = 4000, max_cll: int = 4000, max_fall: int = 400) -> None:
    """SHDMI4 → SIF4 → SHDR4 → EXPDN4 순서로 HDR10(ST2084) 전환.

    on=True  : Dynamic Range On, EOTF=2 (SMPTE ST2084)
    on=False : Dynamic Range Off, EOTF=0 (SDR Range)

    SHDR4 [20H C5H] 파라미터 (section 2.140):
      Program NO=0, On/Off, Type=7, Version=1, EOTF,
      MetaID=0, Disp Primaries x0-y2, White Point x/y,
      Max/Min Disp Mastering, Content/Frame-ave Light LV, Data Type=0
    """
    print(f"\n  {'HDR10 ON (ST2084)' if on else 'SDR 복귀'} 설정 중...")

    # Step 1: 4K60p 타이밍 로드
    send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0) 타이밍 로드")
    time.sleep(1.0)

    if on:
        # Step 2: SHDMI4 [20H 36H] — YCbCr422, 10-bit, BT.2020
        send(build(0x20, 0x36, 0, 1, 2, 0, 2, 0, 0), "SHDMI4 (YCbCr422,10bit,BT.2020)")
        time.sleep(0.3)

        # Step 3: SIF4 [20H 38H] — AVI InfoFrame BT.2020 colorimetry
        send(build(
            0x20, 0x38,
            0, 1, 0, 0, 0,
            2, 2, 0, 0, 0, 1, 0, 2, 3, 119, 1, 0, 0, 0, 0, 0, 0, 6, 0,
            3, 1, "        ", "                ", 0,
            4, 1, 0, 0, 0, 0, 0, 0, 0,
            5, 1, 0, 0, 0,
            0, 0, 0, 0, 0, 0, 0, 0, 0,
            "00000000", "00000000", "0000",
            "00000000", "00000000", "00000000", "00000000",
        ), "SIF4 AVI InfoFrame BT.2020")
        time.sleep(0.1)

        # Step 4: SHDR4 [20H C5H] — HDR10 메타데이터 ON, EOTF=ST2084(2)
        max_m = max(1, int(max_lum * 10000))
        min_m = 1  # 0.0001 nit * 10000 = 1
        send(build(
            0x20, 0xC5,
            0,          # Program NO
            1,          # On/Off = ON
            7,          # Type
            1,          # Version
            2,          # EOTF = SMPTE ST2084
            0,          # Metadata ID
            70800, 29200, 17000, 79700, 13100, 4600,  # Disp Primaries (BT.2020 * 100000)
            31270, 32900,                              # White Point D65 * 100000
            max_m, min_m, max_cll, max_fall,
            0,          # Data Type = HDMI
        ), "SHDR4 HDR10 ON (EOTF=ST2084)")
        time.sleep(0.1)
    else:
        # HDR OFF: SHDR4 On/Off=0, EOTF=0 (SDR)
        send(build(
            0x20, 0xC5,
            0, 0, 7, 1,
            0,  # EOTF = SDR Range
            0,
            70800, 29200, 17000, 79700, 13100, 4600,
            31270, 32900,
            40000000, 1, 4000, 400,
            0,
        ), "SHDR4 HDR OFF (EOTF=SDR)")
        time.sleep(0.1)

    # Step 5: EXPDN4(0,0) — 설정 반영
    send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0) 설정 반영")
    time.sleep(0.5)


run_hdr = input("\nHDR10 (ST2084) 출력 테스트 실행? (y/N) ").strip().lower() == 'y'
if run_hdr:
    print("\n=== Test 4-A: HDR10 ON — 10% APL 흰색 윈도우 ===")
    set_hdr10(on=True)

    # HDR 모드: 8-bit → 10-bit 스케일 (255*4+3=1023)
    x1, y1, x2, y2 = apl_window(10.0)
    send(build(0x28, 0x60), "ALLCLR4")
    time.sleep(0.3)
    send(build(0x28, 0x71, 0, 0, 0, 10), "BCOL4 black 10bit")
    time.sleep(0.2)
    send(build(0x28, 0x61, x1, y1, x2, y2), f"WINDOW4 10%APL ({x1},{y1},{x2},{y2})")
    time.sleep(0.2)
    send(build(0x28, 0x62, 1023, 1023, 1023, 10), "WINCOL4 white 10bit")
    time.sleep(0.2)
    send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
    time.sleep(1.0)
    input("HDR10 10%APL 흰 박스 보임? TV에서 HDR 아이콘 확인. (Enter) ")

    print("\n=== Test 4-B: HDR10 ON — 풀필드 흰색 ===")
    send(build(0x28, 0x60), "ALLCLR4")
    time.sleep(0.3)
    send(build(0x28, 0x61, 0, 0, 3839, 2159), "WINDOW4 전체")
    time.sleep(0.2)
    send(build(0x28, 0x62, 1023, 1023, 1023, 10), "WINCOL4 white 10bit")
    time.sleep(0.2)
    send(build(0x24, 0x20, 0, 0), "EXPDN4(0,0)")
    time.sleep(1.0)
    input("HDR10 풀필드 흰색 OK? (Enter) ")

    print("\n=== Test 4-C: SDR 복귀 ===")
    set_hdr10(on=False)
    send(build(0x24, 0x20, 2286, 0), "EXPDN4(2286,0) 컬러바 로드")
    time.sleep(1.5)
    input("SDR 컬러바 복귀 OK? HDR 아이콘 사라짐? (Enter) ")

# ---------------------------------------------------------------------------
# Restore colourbar
# ---------------------------------------------------------------------------
print("\n컬러바로 복귀...")
send(EXP2286, "EXP2286 복귀")
s.close()
print("완료.")
