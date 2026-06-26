# 패널 측정 프로그램 (LED Measure)

CA-310 / CA-410 색채휘도계와 VG-876 / VG-879 패턴 제너레이터를 연동하여  
디스플레이 패널의 광학 측정 시퀀스를 자동화하는 데스크탑(PySide6) + 웹(Flask) 프로그램입니다.

---

## 목차

1. [기술 스택](#기술-스택)
2. [프로젝트 구조](#프로젝트-구조)
3. [아키텍처 원칙](#아키텍처-원칙)
4. [데이터 모델](#데이터-모델)
5. [장비 계층 (Equipment)](#장비-계층-equipment)
6. [시퀀스 계층 (Sequences)](#시퀀스-계층-sequences)
7. [엔진 계층 (Engine)](#엔진-계층-engine)
8. [UI 계층 — 데스크탑](#ui-계층--데스크탑)
9. [UI 계층 — 웹](#ui-계층--웹)
10. [Excel 내보내기 (Export)](#excel-내보내기-export)
11. [보고서 템플릿](#보고서-템플릿)
12. [명암비 기준](#명암비-기준)
13. [스레딩 모델](#스레딩-모델)
14. [설치 및 실행](#설치-및-실행)
15. [웹 버전 서버 배포](#웹-버전-서버-배포)
16. [전체 자동화 측정 (All-in-One)](#전체-자동화-측정-all-in-one)
17. [EXE 빌드 (Windows)](#exe-빌드-windows)
16. [Windows 포트 연결 설정](#windows-포트-연결-설정)
17. [장비 통신 사양 (CA-310 USB / CA-410 RS-232)](#장비-통신-사양)

---

## 기술 스택

| 역할 | 라이브러리 |
|------|-----------|
| 데스크탑 UI | PySide6 (Qt 6) |
| 웹 UI | Flask + Jinja2 + SSE |
| 시리얼 통신 | pyserial |
| Excel 내보내기 | openpyxl |
| PPT 내보내기 | python-pptx + lxml |
| 웹 PPT 차트 렌더링 | matplotlib |
| EXE 패키징 | PyInstaller |
| Python 버전 | 3.10 이상 |

---

## 프로젝트 구조

```
led_measure/
├── core/                          ← UI 독립 공통 엔진 (재사용 가능)
│   ├── engine.py                  MeasurementEngine — 시퀀스 오케스트레이터
│   ├── export.py                  Excel(.xlsx) 내보내기 전담 (ExcelExporter 클래스)
│   ├── colorimetry.py             Robertson 보간법 기반 CCT / Duv 계산
│   ├── gamut_utils.py             u'v' 색역 면적 및 DCI/BT.2020 overlap 계산
│   ├── equipment/
│   │   ├── base.py                MeterBase, GeneratorBase ABC + 공유 데이터클래스
│   │   ├── ca_meter.py            CA-310 / CA-410 pyserial 드라이버
│   │   ├── vg_generator.py        VG-876 / VG-879 패턴 제너레이터 드라이버
│   │   └── mock.py                하드웨어 없이 UI 테스트용 Mock 장비
│   └── sequences/
│       ├── center_align.py        1. 센터 맞추기
│       ├── lum_swing.py           2. 휘도 스윙 (연속 측정 스트림)
│       ├── lum_loading.py         3. APL 로딩
│       ├── gamut.py               4. 색재현율 (R→G→B→W→BK)
│       └── contrast.py            5. 명암비 (Full White → Black 윈도우 크기별)
│
├── desktop/                       ← PySide6 데스크탑 UI
│   ├── main_window.py             MainWindow + 6개 시퀀스 패널 전부 포함
│   └── worker.py                  MeasurementWorker (QThread 래퍼)
│
├── web/                           ← Flask 웹 UI
│   ├── app.py                     Flask 애플리케이션 팩토리
│   ├── routes.py                  REST API + SSE 진행상황 스트림
│   └── templates/index.html       단일 페이지 UI (Canvas 차트 포함)
│
├── run_desktop.py                 데스크탑 실행 진입점
├── led_measure.spec               PyInstaller 빌드 설정
├── build_windows.bat              Windows EXE 빌드 스크립트
└── requirements.txt
```

---

## 아키텍처 원칙

```
┌──────────────────────────────────────┐   ┌──────────────────────────────┐
│         PySide6 Desktop              │   │        Flask Web             │
│  main_window.py  worker.py(QThread)  │   │  routes.py  index.html(SSE)  │
└────────────────────┬─────────────────┘   └──────────────┬───────────────┘
                     │ import                              │ import
              ┌──────▼──────────────────────────────────── ▼──────┐
              │                  core/                            │
              │  MeasurementEngine                                │
              │  ├── meter: MeterBase         (CA-410 or Mock)    │
              │  ├── generator: GeneratorBase (VG-879 or Mock)    │
              │  ├── meter_lock: threading.Lock                   │
              │  ├── on_progress: Callable[[step,pct,data],None]  │
              │  ├── session_swing / loading / gamut / contrast   │
              │  └── run_sequence(name, **kwargs) → dict          │
              │                                                   │
              │  Sequences         Equipment        Export        │
              │  ├─ CenterAlign    ├─ CaMeter       ExcelExporter │
              │  ├─ LumSwing       ├─ VgGenerator                 │
              │  ├─ LumLoading     └─ Mock                        │
              │  ├─ Gamut                                         │
              │  └─ Contrast                                      │
              └──────────────────────────┬────────────────────────┘
                                         │ pyserial
              ┌──────────────────────────▼────────────────────────┐
              │                 Hardware                          │
              │    CA-310 / CA-410          VG-876 / VG-879       │
              └───────────────────────────────────────────────────┘
```

**핵심 규칙**:
- `core/` 는 `desktop/` 또는 `web/` 을 절대 import하지 않습니다.
- UI가 바뀌어도 측정 로직은 그대로 재사용됩니다.
- `on_progress` 콜백만 교체하면 데스크탑·웹·CLI 모두 동일한 엔진을 구동합니다.

---

## 데이터 모델

### MeasureResult — 단일 측정값 단위

```python
@dataclass
class MeasureResult:
    timestamp_ms: int      # 측정 시각 (Unix ms)
    Lv:           float    # 휘도 (cd/m²)
    x:            float    # CIE x 색도
    y:            float    # CIE y 색도
    u_prime:      float    # CIE u' 색도
    v_prime:      float    # CIE v' 색도
    X:            float    # CIE X 삼자극값
    Y:            float    # CIE Y (= Lv)
    Z:            float    # CIE Z 삼자극값
    cct:          float    # 상관 색온도 (K)
    duv:          float    # Planckian locus 거리
    pattern_info: PatternInfo   # 측정 당시 패턴 메타데이터
```

### PatternInfo — 패턴 메타데이터 (MeasureResult에 포함)

```python
@dataclass
class PatternInfo:
    type:       str    # "full_field" | "window" | "raster_window" | "crosshair"
    apl_pct:    float  # 화면 평균 밝기 %
    width_pct:  float  # 창 너비 %
    height_pct: float  # 창 높이 %
    color:      str    # "white" | "black" | "red" | "green" | "blue"
    is_hdr:     bool   # HDR 출력 여부
```

### PatternConfig — 장비에 보낼 패턴 명령 (시퀀스→제너레이터)

```python
@dataclass
class PatternConfig:
    type:       str    # 위와 동일
    color:      str
    r: int; g: int; b: int           # 0~255 (8-bit) 또는 0~1023 (10-bit)
    width_pct:  float = 100.0
    height_pct: float = 100.0
    bg_r: int = 0; bg_g: int = 0; bg_b: int = 0
    bit_mode:   int = 8
```

---

## 장비 계층 (Equipment)

### 공통 인터페이스 (base.py)

```python
class MeterBase(ABC):
    def connect(self, port: str) -> None: ...
    def disconnect(self) -> None: ...
    def measure(self) -> MeasureResult: ...
    def set_current_pattern(self, info: PatternInfo) -> None: ...
    @property
    def is_connected(self) -> bool: ...

class GeneratorBase(ABC):
    def connect(self, port: str) -> None: ...
    def disconnect(self) -> None: ...
    def set_pattern(self, cfg: PatternConfig) -> None: ...
    def set_hdr(self, enabled: bool) -> None: ...
    def set_sdr(self) -> None: ...
    def reset(self) -> None: ...
    @property
    def is_connected(self) -> bool: ...
```

### CA-410 / CA-310 (ca_meter.py)

---

#### CA-310 — USB 직접 연결 (네이티브 USB, 기본 방식)

CA-310은 USB-Serial 변환 칩을 사용하지 않는 **네이티브 USB 장치**입니다.  
`/dev/ttyUSB*` 또는 `COM*` 포트에 잡히지 않으며, `lsusb` 에 VID=0686:1002 로 노출됩니다.

**USB 장치 정보**:

| 항목 | 값 |
|------|-----|
| Vendor ID (VID) | `0x0686` (Minolta) |
| Product ID (PID) | `0x1002` |
| USB 클래스 | Vendor-specific (0xFF) |
| USB 버전 | USB 1.1 |
| 통신 방식 | Bulk Transfer |

**사용 엔드포인트**:

| 방향 | Endpoint | 패킷 크기 | 용도 |
|------|----------|----------|------|
| OUT (명령 전송) | `0x02` | 64 byte | 호스트 → 장치 |
| IN (응답 수신) | `0x82` | 64 byte | 장치 → 호스트 |

> EP1 (0x01 / 0x81, 16-byte)은 통신에 사용되지 않음 — 모든 명령·응답은 EP2 경유.

**Linux udev 설정** (non-root 접근):
```bash
echo 'SUBSYSTEM=="usb", ATTR{idVendor}=="0686", ATTR{idProduct}=="1002", MODE="0666", GROUP="plugdev"' \
  | sudo tee /etc/udev/rules.d/99-ca310.rules
sudo udevadm control --reload-rules && sudo udevadm trigger
```

**CA-310 USB 연결 시퀀스**:

```
1. usb.core.find(idVendor=0x0686, idProduct=0x1002)  ← 장치 탐색
2. dev.reset()                                         ← USB 리셋
   time.sleep(1.5)                                     ← 재열거 대기
   dev = usb.core.find(...)                            ← 재탐색
3. dev.detach_kernel_driver(0)  (커널 드라이버 있으면)
4. dev.set_configuration()
   time.sleep(0.5)
5. EP2 OUT 전송: b"COM,1\r\n" + b"\x00" * 57  (64-byte 패딩 필수)
   EP2 IN 수신:  "OK00..."  ← REMOTE 모드 전환 확인
6. EP2 OUT 전송: b"MDS,0\r\n" + b"\x00" * 57
   EP2 IN 수신:  "OK00..."  ← 측정 모드 CIE1931 xyLv 고정
```

> **64바이트 패딩 필수**: CA-310은 USB 패킷이 64바이트 미만이면 응답하지 않습니다.  
> `dev.reset()` 생략 시 cold-start 상태에서 첫 명령을 무시하는 동작이 있어 reset 필수.

**명령 응답 형식**:
- 성공: `OK00` + 공백 패딩 (총 64바이트 이상)
- 오류: `ER10` + 공백 패딩
- `IDQ` 명령은 CA-310에서 지원하지 않음 → `ER10` 반환

**MES 응답 형식**:
```
OK00,P1 2665;4228;.8758
         ↑     ↑     ↑
         x×10000  y×10000  Lv(cd/m²)
→ x=0.2665, y=0.4228, Lv=0.8758
```

**해제 시퀀스**:
```python
dev.write(0x02, b"COM,0\r\n" + b"\x00" * 57, timeout=1000)
usb.util.dispose_resources(dev)
```

---

#### CA-310 — RS-232 시리얼 (폴백 / 별도 RS-232 어댑터 사용 시)

| 파라미터 | 값 |
|---------|-----|
| Baud Rate | 38400 |
| Data Bits | 7 |
| Parity | EVEN |
| Stop Bits | 2 |
| Flow Control | 없음 (RTS/CTS 미사용) |
| Timeout | 5.0 s |

**연결 시퀀스**:
1. `serial.Serial(port, 38400, 7, EVEN, 2, rtscts=False)` 포트 열기
2. `COM,1\r\n` → REMOTE 모드 전환 (응답에 "OK" 포함 확인)
3. `IDQ\r\n` → 기기 ID 조회

> CA-310을 USB 연결할 때 RS-232 경로를 사용할 필요 없음.  
> RS-232는 별도 어댑터를 통해 장치 후면 DB-9 포트를 직접 연결하는 경우에만 해당.

---

#### CA-410 — RS-232 시리얼

| 파라미터 | 값 |
|---------|-----|
| Baud Rate | 38400 |
| Data Bits | 7 |
| Parity | EVEN |
| Stop Bits | 2 |
| Flow Control | RTS/CTS (필수) |
| Timeout | 5.0 s |

**연결 시퀀스**:
1. `serial.Serial(port, 38400, 7, EVEN, 2, rtscts=True)` 포트 열기
2. `COM,1\r\n` → REMOTE 모드 전환 (응답에 "OK" 포함 확인)
3. `MDR,0\r\n` → 홀드 해제
4. `IDQ\r\n` → 기기 ID 조회

> CA-410은 **RTS/CTS 필수**. CH340 계열 저가 USB-Serial은 미지원 → FTDI FT232R 또는 CP210x 권장.  
> RS-232 케이블은 **스트레이트 케이블** 사용 (Null-Modem/크로스 금지).

---

**측정 (CA-310 USB / CA-410 공통)**:
- 명령: `MES\r\n` (USB는 64바이트 패딩)
- 응답 형식 1 (표준): `OK,Lv,x,y,u',v',CCT,Duv\r`
- 응답 형식 2 (Compact / CA-310 USB): `OK00,Pxx v1;v2;v3`  (v1=x×10000, v2=y×10000, v3=Lv)
- 저휘도(< 0.01 cd/m²) 시 응답 지연 → `timeout=10.0s`

**해제**: `COM,0\r\n` → LOCAL 복귀 → 포트/USB 닫기

### VG-876 / VG-879 (vg_generator.py)

**시리얼 파라미터**: 38400 baud / 8N1 / 흐름 제어 없음

**바이너리 프레임 구조**:
```
STX(02H) + FDH + CMD1 + CMD2 + ASCII파라미터(콤마구분) + ETX(03H)
```

**주요 커맨드**:

| 커맨드 | 코드 | 설명 |
|--------|------|------|
| EXPDN4 | `24H 20H` | 프로그램 실행 / 타이밍 로드 (`EXPDN4(0,0)` = 현재 설정 렌더링) |
| ALLCLR4 | `28H 60H` | 모든 플레인(윈도우+베이스) 클리어 |
| WINDOW4 | `28H 61H` | 윈도우 좌표 등록 (H/V 라인 메모리에 누적) |
| WINCOL4 | `28H 62H` | 등록된 모든 창에 색상 적용 |
| WINCLR4 | `28H 63H` | 윈도우 플레인만 클리어 (베이스 패턴 유지) |
| BCOL4 | `28H 71H` | 창 바깥 배경색 |
| SHDMI4 | `20H 36H` | HDMI 출력 포맷 (색공간/비트심도) |
| SIF4 | `20H 38H` | AVI InfoFrame (BT.2020 색역 선언) |
| SHDR4 | `20H C5H` | HDR10 Dynamic Range 메타데이터 |
| SPT4 | `20H 2CH` | 프리셋 패턴 색상 설정 |
| SPTS4 | `20H 2AH` | 프리셋 패턴 선택 + 실행 |

**패턴 출력 Python API**:

```python
# 전체화면 단색 (색재현율용)
show_full_field(r, g, b, bit_mode=8)
# → 처음: ALLCLR4 → WINDOW4(전체) → WINCOL4 → EXPDN4
# → 연속: WINCOL4만으로 즉시 색상 전환 (깜빡임 없음)

# 중앙 윈도우 패치 (APL 로딩용)
show_window_patch(width_pct, height_pct, r, g, b, bg_r=0, bg_g=0, bg_b=0, bit_mode=8)
# → ALLCLR4 → WINDOW4(중앙) → WINCOL4(창색) → EXPDN4
# → 밝은 배경 시: 4개 스트립으로 배경 등록 후 WINCOL4 1회

# 흰색 래스터 + 중앙 검은 창 (명암비용)
show_white_raster_black_window(side_pct, bit_mode=8)
# → ALLCLR4 → SPT4(fg=white) → SPT4(bg=black) → SPTS4(0,1,2,10) → EXPDN4(9999,0)
# → WINDOW4(중앙 black 창) → WINCOL4(black) → EXPDN4

# HDR10 전환
set_hdr(enabled=True)
# → EXPDN4(2286,0) 타이밍 로드 → SHDMI4 → SIF4 → SHDR4 → EXPDN4

# SDR 전환
set_sdr()
# → EXPDN4(2286,0) → SHDR4(Off) → EXPDN4

# 장비 동결 시 복구
reset()
# → ENQ 터미널 재진입 → EXPDN4(2286,0) 컬러바 복귀
```

> **`_reset_window_memory()` 내부 동작**: ALLCLR4 [28H 60H] 사용.  
> WINCLR4 [28H 63H]는 컬러바 베이스를 남겨 BCOL4 배경색이 무효화되는 버그 있음.  
> ALLCLR4는 모든 플레인(윈도우+베이스)을 완전 소거하므로 이후 BCOL4 정상 동작.

**SDR 초기화 한 번만**: `_timing_loaded` 플래그로 관리.  
`set_sdr()` / `reset()` 호출 시 `_timing_loaded = False` → 다음 패턴 출력 시 `EXPDN4(2286,0)` 1회 재로드.  
이후 패턴 전환에서는 재로드 생략 → gamut R→G→B 전환 시 컬러바 깜빡임 없음.

### Mock 장비 (mock.py)

실제 장비 없이 UI 동작 테스트용.  
`MockMeter.measure()`는 `pattern_info`의 색상과 APL에 따라 현실적인 Lv/xy/u'v'/XYZ/CCT/Duv 값을 반환.  
`MockGenerator`는 상태만 추적하고 실제 시리얼 출력 없음.

---

## 시퀀스 계층 (Sequences)

각 시퀀스 클래스는 `__init__(engine)` + `run(**kwargs) → dict` + `stop()` 구조.  
`engine.generator`, `engine.meter`, `engine.meter_lock`에 직접 접근.

### 1. 센터 맞추기 (center_align.py)

```
VG: 흰색 1000×1000 px 박스를 3840×2160 화면 정중앙에 출력
    (SPTS4 + SPTS4 조합: ABC 버튼 + ㅁ + X + R/G/B 동시 활성)
→ 사용자가 CA 렌즈를 화면 정중앙에 맞춤
→ UI에서 [OK] 클릭으로 완료
```

### 2. 휘도 스윙 (lum_swing.py)

```
패턴: Black 바탕 + White 윈도우 (W=31.6%, H=31.6% → APL 10%)
      show_window_patch(31.6, 31.6, 255, 255, 255, 0, 0, 0)
측정: 301회 연속, 1 sample/sec (절대시각 기반 스케줄링)
출력: List[MeasureResult] (시간 순서)
```

### 3. APL 로딩 (lum_loading.py)

**APL 스텝 정의**:

| 버전 | 스텝 수 | 포인트 |
|------|---------|--------|
| 37단계 | 37 | 1~100 구간 비선형 (1,2,3,4,5,6,8,9,10,11,12,14,16,18,20,23,25,28,30,33,36,39,42,46,49,53,56,60,64,68,72,77,81,86,90,95,100) |
| 10단계 | 10 | 1,3,10,16,25,36,49,64,81,100 |
| 2단계  |  2 | 10, 100 |

**APL → 창 크기 변환**: `side = sqrt(apl/100) × 100 %`  
(예: APL 10% → side = 31.62%)

**측정 파라미터**:
- `measurements_per_step` (기본 3회): 스텝당 측정 횟수
- `_INTER_MEAS_SLEEP` (기본 0.3s): 동일 스텝 내 측정 간격 (디스플레이 안정 후 미터 버퍼 소거 목적)
- `cooling_enabled`: APL ≤ threshold 시 `show_black()` 후 cooling_duration_sec 대기
- `cooling_apl_threshold` (기본 10): 쿨링 적용 APL 임계값
- `cooling_duration_sec` (기본 5.0): 쿨링 시간

**출력**: `Dict[int, List[MeasureResult]]` — `{apl_pct: [r1, r2, r3]}`

**progress 콜백 데이터**: `{"apl": apl, "results": [MeasureResult, ...]}`

### 4. 색재현율 (gamut.py)

```
측정 순서: Red → Green → Blue → White → Black (각 1회)
패턴: show_full_field(r, g, b)  (100% 전체 화면)
출력: Dict[str, MeasureResult]  {"red": r, "green": r, ...}
```

**DCI-P3 / BT.2020 Overlap 계산** (`gamut_utils.py`):  
- R/G/B의 u'v' 좌표로 측정 삼각형 생성
- Sutherland-Hodgman 알고리즘으로 기준 삼각형과의 교차 면적 계산
- overlap(%) = 교차 면적 / 기준 삼각형 면적 × 100

**BT.2020 기준점 (u', v')**: `[(0.197,0.468), (0.092,0.018), (0.508,0.156)]`  
**DCI-P3 기준점 (u', v')**: `[(0.209,0.488), (0.139,0.050), (0.460,0.146)]`

### 5. 명암비 (contrast.py)

```
측정 순서:
  Step 0 → Full White (full_field, 255,255,255)  ← 기준 Lv
  Step 1 → Black 창 100×100%  (raster_window)
  Step 2 → Black 창 50×50%
  Step 3 → Black 창 20×20%
  Step 4 → Black 창 14.1×14.1%

CR = Full_White_Lv / Black_Window_Lv  (각 스텝별)

출력: Dict[float, MeasureResult]  {0.0: white_r, 100.0: r, 50.0: r, 20.0: r, 14.1: r}
```

**progress 콜백 데이터**: `{"win_size": 0.0 또는 100.0/50.0/20.0/14.1, "result": MeasureResult}`

---

## 엔진 계층 (Engine)

### MeasurementEngine (engine.py)

```python
class MeasurementEngine:
    brand: str
    model_name: str
    auto_save_dir: str

    meter: MeterBase | None
    generator: GeneratorBase | None
    meter_lock: threading.Lock      # 측정 중 동시 접근 방지

    on_progress: Callable[[str, float, Any], None]  # UI가 교체하는 콜백

    # 세션 누적 저장 (측정 완료 시 자동으로 _all.xlsx 재저장)
    session_swing:    Dict[str, List[MeasureResult]]          # "SDR_Vivid" → [r...]
    session_loading:  Dict[str, Dict[int, List[MeasureResult]]] # "SDR_Vivid" → {apl→[r...]}
    session_gamut:    Dict[str, Dict[str, MeasureResult]]     # "SDR" → {"red"→r, ...}
    session_contrast: Dict[str, Dict[float, MeasureResult]]   # "SDR" → {0.0→r, 100.0→r, ...}
```

**`run_sequence(name, **kwargs)`**: 시퀀스 이름으로 해당 `_run_*` 메서드 디스패치.  
내부적으로 `on_progress`를 호출하므로 UI·웹 양쪽 모두 동일하게 동작.

**`on_progress(step, pct, data)` 콜백 규약**:

| step | pct | data |
|------|-----|------|
| `"lum_swing"` | 0.0~1.0 | `MeasureResult` 또는 `List[MeasureResult]` |
| `"lum_loading"` | 0.0~1.0 | `{"apl": int, "results": [MeasureResult]}` |
| `"gamut"` | 0.0~1.0 | `{"color": str, "result": MeasureResult}` |
| `"contrast"` | 0.0~1.0 | `{"win_size": float, "result": MeasureResult}` |

**`connect_meter(port, model)`** / **`connect_generator(port, model)`**: 적절한 드라이버 인스턴스 생성 후 연결.  
**`disconnect_all()`**: 모든 장비 해제.  
**`is_ready`**: `meter.is_connected and generator.is_connected`.

---

## UI 계층 — 데스크탑

### 구조 (main_window.py)

```
MainWindow (QMainWindow)
├── ConnectionPanel       장비 연결 / 포트 스캔 / Brand·Model 입력 / 자동저장 폴더
├── CenterAlignPanel      센터 패턴 출력 버튼
├── LumSwingPanel         케이스·HDR 선택 → 측정 → Canvas 라인 차트 + 테이블
├── LumLoadingPanel       버전·케이스·HDR·쿨링 설정 → 측정 → SDR/HDR 차트 + 테이블
├── GamutPanel            HDR 선택 → 측정 → u'v' 차트 + 테이블
├── ContrastPanel         HDR 선택 → 측정 → 테이블 (Full White + Black 창 4종)
└── ReportPanel           xlsx 파일 불러오기 → 모델 비교 표 + APL/u'v' 차트
```

### 스레딩 패턴 (worker.py)

```python
class MeasurementWorker(QThread):
    progress = Signal(str, float, object)   # step, pct, data
    succeeded = Signal(object)              # 최종 결과 dict
    error = Signal(str)                     # 예외 메시지

    def run(self):
        # engine.on_progress를 progress.emit으로 교체
        # engine.run_sequence() 실행
        # 완료 시 succeeded.emit(result)
```

**패널 → 워커 연결 패턴**:
```python
worker = MeasurementWorker(engine, "contrast", is_hdr=True)
worker.progress.connect(self._on_progress)   # QueuedConnection (자동)
worker.succeeded.connect(self._on_finished)
wire_worker_cleanup(worker, self, '_worker') # QThread::finished → wait() + setattr None
worker.start()
```

**실시간 업데이트 원칙**: `_on_progress`에서 **행/포인트 추가만** 수행. 전체 테이블 재구성(`setRowCount(0)`) 또는 시리즈 삭제(`removeAllSeries()`) 금지 → O(n²) 방지.

### ReportPanel 동작 흐름

1. `_load_files()` → xlsx 선택 → `_parse_xlsx(path)`
2. `_parse_xlsx` → Info 시트에서 Brand/Model/Sequence 파악 → 적절한 파서 호출:
   - `_parse_lum_loading_wb()`: Raw_ 또는 Summary 시트에서 APL별 Lv 추출
   - `_parse_gamut_wb()`: Gamut 시트에서 u'v' 좌표 추출 → DCI/BT.2020 overlap 계산
   - `_parse_all_sessions_wb()`: Loading_*/Gamut_*/Contrast_* 시트 일괄 파싱
3. `_find_or_create_entry(brand, model)` → 모델별 데이터 dict 관리
4. `_refresh_report_table()` → 현재 선택된 형식으로 표 렌더링

**보고서 형식 선택**:
- **경쟁사 비교 장표**: HDR 10%, HDR 100%, SDR 10%, SDR 100%, CR, Black Lv, DCI, BT.2020
- **광학 측정 데이터**: Vivid SDR 10%/100%, Standard SDR 10%/100%, Vivid HDR 10%/100%, Standard HDR 10%/100%, Cinema HDR 10%/100%, CR, DCI, BT.2020

**entry 데이터 구조**:
```python
entry = {
    "brand": str, "model": str,
    # 경쟁사 비교용
    "hdr_10": float, "hdr_100": float,
    "sdr_10": float, "sdr_100": float,
    "contrast_ratio": float,   # Full White Lv / Black 100% Lv
    "black_lv": float,         # Black 100% 창 Lv
    "dci_overlap": float,      # DCI-P3 overlap %
    "bt2020_overlap": float,   # BT.2020 overlap %
    # 광학 측정용 (케이스별)
    "sdr_vivid_10": float, "sdr_vivid_100": float,
    "sdr_standard_10": float, "sdr_standard_100": float,
    "hdr_vivid_10": float, "hdr_vivid_100": float,
    "hdr_standard_10": float, "hdr_standard_100": float,
    "hdr_cinema_10": float, "hdr_cinema_100": float,
    # 차트용
    "apl_sdr": Dict[int, float],   # {apl: avg_lv}
    "apl_hdr": Dict[int, float],
    "gamut_uv": Dict[str, tuple],  # {"red": (u', v'), ...}
}
```

---

## UI 계층 — 웹

### 구조 (web/)

```
app.py          Flask 팩토리 — engine을 app.extensions["engine"]에 저장
routes.py       Blueprint "api" (/api/*) + Blueprint "ui" (/)
index.html      단일 페이지 — 사이드바 + 패널 전환 + Canvas 차트
```

**API 엔드포인트**:

| 메서드 | 경로 | 설명 |
|--------|------|------|
| GET | `/api/ports` | 시리얼 포트 목록 |
| POST | `/api/connect` | 장비 연결 (brand, model_name, meter_port, gen_port 등) |
| POST | `/api/mock_connect` | Mock 장비 연결 |
| POST | `/api/disconnect` | 전체 해제 |
| GET | `/api/status` | 연결 상태 |
| POST | `/api/hdr` | HDR/SDR 전환 `{"enabled": bool}` |
| POST | `/api/run/<seq_name>` | 시퀀스 시작 (백그라운드 스레드) |
| POST | `/api/stop` | 시퀀스 중단 |
| GET | `/api/progress` | SSE 진행상황 스트림 |
| GET | `/api/export/<type>` | Excel 다운로드 (`lum_swing` / `lum_loading` / `gamut` / `contrast`) |
| GET | `/api/export/ppt` | PPT 보고서 다운로드 (현재 세션 데이터 기반) |

**SSE 이벤트 형식**:
```json
{"step": "contrast", "progress": 0.4, "data": {"win_size": 100.0, "result": {...}}}
{"step": "contrast", "progress": 1.0, "done": true}
{"error": "오류 메시지"}
```

**웹 차트**: Canvas 2D API 기반 직접 구현 (외부 CDN 의존 없음).

**PPT 보고서 (`/api/export/ppt`)**:  
- 현재 세션의 `session_loading` / `session_gamut` / `session_contrast` 데이터를 사용  
- APL 로딩 (SDR/HDR), u'v' 색도 차트는 matplotlib으로 서버에서 렌더링하여 슬라이드에 삽입  
- 테이블 테두리는 Windows PowerPoint 호환 처리 적용 (`prstDash val="solid"` + `tableStyleId` 제거)  
- 차트 글자 크기: 타이틀 11pt, 범례·축 10pt  
- 추가 의존성: `pip install python-pptx lxml matplotlib`

---

## Excel 내보내기 (Export)

모든 Excel 양식은 `core/export.py` 의 `ExcelExporter` 클래스 한 곳에서 정의.

### 생성 시트 목록

| 함수 | 생성 시트 |
|------|-----------|
| `export_lum_swing` | `Info`, `Case_Vivid` / `Case_Standard` / `Case_Cinema` |
| `export_lum_loading` | `Info`, `Summary` (APL×케이스 집계), `Raw_Vivid` 등 |
| `export_gamut` | `Info`, `Gamut` |
| `export_contrast` | `Info`, `ContrastRatio` |
| `export_all_session` | `Info`, `Swing_*`, `Loading_Summary`, `Loading_*`, `Gamut_*`, `Contrast_*` |
| `export_report_template` | `Info`, `Summary`, `Gamut`, `LumLoading`, `Contrast` |

### 공통 측정값 컬럼 (`_MEAS_COLS`)

```
Time(s) | Lv(cd/m²) | x | y | u' | v' | X | Y | Z | CCT(K) | Duv | Pattern | APL(%) | W(%) | H(%) | Color | SDR/HDR
```

### `export_all_session` — 자동 저장 통합 파일

`{brand}_{model}_all.xlsx`로 저장. 각 측정 완료 시마다 덮어씀 (누적 갱신).

**Contrast 시트 컬럼**: `Black H/V (%) | Lv (cd/m²) | CR (White/Lv) | ...측정값 컬럼...`  
- `Black H/V (%)` = "Full White" 또는 100.0 / 50.0 / 20.0 / 14.1
- `CR` = Full White Lv / Black Window Lv (Full White 행은 "—")

---

## 보고서 템플릿

### 데이터 파싱 흐름

```
xlsx 파일 업로드
    │
    ├── Info 시트 → Brand, Model, Sequence 확인
    │
    ├── Sequence == "All Sessions"
    │       → Loading_* 시트: APL별 Lv 집계 (최대/중간/최소)
    │       → Gamut_* 시트: u'v' 좌표 → DCI/BT.2020 overlap 계산
    │       └── Contrast_* 시트: Full White Lv + Black 100% Lv → CR 계산
    │
    ├── Sequence == "Luminance Loading"
    │       → Raw_* 또는 Summary 시트에서 APL별 Lv 추출
    │
    └── Sequence == "Gamut"
            → Gamut 시트에서 u'v' 좌표 추출
```

### 파일명 → 케이스 모드 자동 감지

- 파일명 또는 시트명에 `"HDR"` 포함 → HDR
- `"Cinema"` 포함 → cinema, `"Standard"` 또는 `"STD"` → standard, 그 외 → vivid

---

## 명암비 기준

**CR 공식**: `CR = Full_White_Lv / Black_Window_Lv`

| Black Window 크기 | 의미 | CR |
|---|---|---|
| **100×100% (전체 검정)** | Full On/Off CR — 업계 홍보용 | 가장 높음 ✅ 장표 기준 |
| 50×50% | 실사용 기준 | 중간 |
| 20×20% | IEC 61947-2 표준 권장 | 낮음 |
| 14.1×14.1% | ICDM 10% APL 기준 | 가장 낮음 |

> 보고서 템플릿 `contrast_ratio` 필드 = `Full White Lv / Black 100% window Lv`  
> 경쟁사 비교 장표에서 가장 높게 표시되는 수치이므로 이 값을 기준으로 사용.

---

## 스레딩 모델

```
Main Thread (Qt Event Loop)
│
├── ConnectionPanel._connect_meter()
│       → ConnectWorker(QThread) — connect() 블로킹 호출
│           → succeeded 시그널 → UI 업데이트
│
└── 각 측정 패널._run()
        → MeasurementWorker(QThread)
            → engine.on_progress = lambda: progress.emit(...)
            → engine.run_sequence(name, **kwargs)
                → Sequence.run()
                    → generator.set_pattern()   (시리얼 I/O)
                    → with meter_lock:
                          meter.measure()       (시리얼 I/O, 블로킹)
                    → on_progress(step, pct, data)
                          → QueuedConnection → Main Thread
                              → _on_progress() : 행 추가 / 포인트 추가
            → succeeded.emit(result) → _on_finished()
```

**`meter_lock`**: 측정 중 다른 스레드(예: HDR 전환 워커)가 미터에 동시 접근하는 것을 막기 위한 `threading.Lock`.

**`wire_worker_cleanup(worker, owner, attr)`**:  
`QThread::finished` 시그널에 `worker.wait() → setattr(owner, attr, None)` 연결.  
Python GC와 Qt C++ 소멸자 타이밍 충돌로 인한 크래시 방지.

---

## 설치 및 실행

### 요구사항

- Python 3.10 이상
- 실제 측정 시 CA-310/CA-410 + VG-876/VG-879 USB-Serial 연결

### 설치

```bash
python -m venv .venv

# Linux/macOS
source .venv/bin/activate

# Windows
.venv\Scripts\activate

pip install -r requirements.txt
```

### 데스크탑 실행

```bash
python run_desktop.py
```

> **Mock 모드**: 장비 없이 테스트 시 연결 패널에서 **[🧪 Mock 연결]** 클릭.

### 웹 실행

```bash
python web/app.py
# → http://localhost:5000
```

---

## 웹 버전 서버 배포

### 필요한 파일 목록

서버에 올릴 때 아래 파일/폴더만 복사하면 됩니다. `desktop/`, `run_desktop.py`, 빌드 파일, 테스트 파일은 불필요합니다.

```
led_measure/
├── run.py                              ← 웹 서버 실행 진입점
├── requirements.txt                    ← 의존성 패키지 목록
│
├── web/                                ← Flask 웹 앱
│   ├── __init__.py
│   ├── app.py                          ← Flask 팩토리
│   ├── routes.py                       ← REST API + SSE 엔드포인트
│   └── templates/
│       └── index.html                  ← 단일 페이지 UI
│
└── core/                               ← 측정 엔진 (UI 독립)
    ├── __init__.py
    ├── engine.py                       ← MeasurementEngine
    ├── export.py                       ← Excel 내보내기
    ├── colorimetry.py                  ← CCT / Duv 계산
    ├── gamut_utils.py                  ← 색역 overlap 계산
    ├── equipment/
    │   ├── __init__.py
    │   ├── base.py                     ← MeterBase / GeneratorBase ABC
    │   ├── ca_meter.py                 ← CA-310 / CA-410 드라이버
    │   ├── vg_generator.py             ← VG-876 / VG-879 드라이버
    │   └── mock.py                     ← 하드웨어 없이 테스트용 Mock
    └── sequences/
        ├── __init__.py
        ├── center_align.py
        ├── lum_swing.py
        ├── lum_loading.py
        ├── gamut.py
        └── contrast.py
```

### 서버 설치 및 실행

```bash
# 1. 가상 환경 생성 및 활성화
python -m venv .venv
source .venv/bin/activate          # Linux / macOS
# .venv\Scripts\activate           # Windows

# 2. 의존성 설치 (웹만 사용 시 PySide6 불필요)
# PPT 내보내기도 사용할 경우 python-pptx lxml matplotlib 추가
pip install flask pyserial openpyxl python-pptx lxml matplotlib

# 3. 실행
python run.py
# → http://<서버IP>:5000
```

> **웹 전용 설치 시** requirements.txt에서 `PySide6` 줄은 설치하지 않아도 됩니다.  
> 실제 장비 없이 기능 확인이 필요하면 브라우저에서 **[🧪 Mock 연결]** 버튼을 사용하세요.

### 서버 운영 환경 (선택사항)

개발용 `flask run` 대신 안정적인 서비스를 위해 Gunicorn + 리버스 프록시 구성을 권장합니다.

```bash
pip install gunicorn

# SSE(Server-Sent Events)는 eventlet / gevent 비동기 워커와 충돌 가능
# 반드시 sync 워커(-k sync) + 스레드 옵션 사용
gunicorn "web.app:create_app()" \
  -w 1 --threads 4 -k sync \
  -b 0.0.0.0:5000
```

> `-w 1` (워커 1개) 고정: 측정 엔진이 프로세스 내 싱글턴으로 동작하므로  
> 멀티 워커로 분기하면 장비 상태가 프로세스마다 달라집니다.

---

## EXE 빌드 (Windows)

### 방법 1 — 배치 스크립트 (가장 간단)

```
build_windows.bat 더블클릭
```

완료 시 `dist\led_measure\led_measure.exe` 생성.

### 방법 2 — 수동 빌드

```bat
python -m venv .venv_win
.venv_win\Scripts\pip install PySide6 openpyxl pyserial pyinstaller
.venv_win\Scripts\pyinstaller led_measure.spec --clean --noconfirm
```

> **주의**: PyInstaller 빌드는 플랫폼 종속적입니다 (Windows 머신에서만 `.exe` 생성).

**`run_desktop.py` — Windows DPI 설정** (EXE에도 적용됨):
```python
os.environ["QT_ENABLE_HIGHDPI_SCALING"] = "1"
os.environ["QT_SCALE_FACTOR_ROUNDING_POLICY"] = "PassThrough"
app.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
window.showMaximized()  # 1500px 고정 대신 최대화로 시작
```

> PyInstaller로 빌드된 EXE는 백신이 오진할 수 있습니다.  
> `dist\led_measure\` 폴더 전체를 백신 예외 경로에 추가하세요.

---

## Windows 포트 연결 설정

### 포트 자동 검색

프로그램 시작 시 `serial.tools.list_ports.comports()`로 자동 스캔.  
Windows에서는 `COM3`, `COM7` 등 COMxx 형식.  
COM10 이상은 pyserial이 자동으로 `\\.\COM10` 처리하므로 별도 조치 불필요.

### USB-Serial 드라이버 (필수)

| 칩셋 | 드라이버 |
|------|---------|
| FTDI FT232R | https://ftdichip.com/drivers/vcp-drivers/ |
| CP2102 / CP2104 | https://www.silabs.com/developers/usb-to-uart-bridge-vcp-drivers |
| CH340 / CH341 | https://www.wch-ic.com/downloads/CH341SER_EXE.html |

드라이버 설치 후 장치 관리자 → **포트(COM & LPT)** 에서 `COMxx` 인식 확인.

### CA-410 주의사항

CA-410은 **RTS/CTS 하드웨어 흐름 제어 필수** (38400 / 7E2 / RTS-CTS).  
- CH340 계열 저가 어댑터 일부는 RTS/CTS 미지원 → **FTDI FT232R** 또는 **CP210x** 권장
- RS-232 케이블: **스트레이트 케이블** 사용 (크로스/Null-Modem 케이블 금지)

CA-310: **네이티브 USB 장치** (VID=0686:1002) — USB-Serial 어댑터 불필요.  
  - Windows: `libusb-win32` 또는 `WinUSB` 드라이버 필요 (Zadig 유틸리티로 설치).  
  - Linux: udev 규칙 설정 또는 sudo 실행.  
  - RS-232 폴백 사용 시 (후면 DB-9): 38400 / 7E2 / 흐름 제어 없음.

### VG-879 / VG-876

38400 / 8N1 / 흐름 제어 없음. 모든 어댑터 정상 동작.

### 연결 순서

1. 장비 전원 ON
2. USB-Serial 어댑터 PC 연결
3. 장치 관리자에서 COMxx 확인
4. 프로그램 실행 → **[포트 스캔]** → 포트 선택 → **[연결]**

> 장비 전원 OFF 상태에서 연결 시도 시: CA-410 `REMOTE 전환 실패`, VG ACK 타임아웃 오류 발생.

---

## 장비 통신 사양

### CA-410 시리얼 파라미터

| 파라미터 | 값 |
|---------|-----|
| Baud Rate | 38400 |
| Data Bits | 7 |
| Parity | EVEN |
| Stop Bits | 2 |
| Flow Control | RTS/CTS |
| Timeout | 5.0 s |

### CA-310 USB 파라미터 (기본 연결 방식)

| 파라미터 | 값 |
|---------|-----|
| 연결 방식 | USB Bulk Transfer (pyusb) |
| VID / PID | 0x0686 / 0x1002 |
| EP OUT (명령) | 0x02 (64-byte bulk) |
| EP IN (응답) | 0x82 (64-byte bulk) |
| 패킷 크기 | 64바이트 패딩 필수 |
| 라이브러리 | pyusb >= 1.2.1 + libusb |

### CA-310 RS-232 파라미터 (폴백 — 후면 DB-9 직접 연결 시)

| 파라미터 | 값 |
|---------|-----|
| Baud Rate | 38400 |
| Data Bits | 7 |
| Parity | EVEN |
| Stop Bits | 2 |
| Flow Control | 없음 |

### VG 시리얼 파라미터

| 파라미터 | 값 |
|---------|-----|
| Baud Rate | 38400 |
| Data Bits | 8 |
| Parity | NONE |
| Stop Bits | 1 |
| Flow Control | 없음 |



---

## 전체 자동화 측정 (All-in-One)

브랜드 타입(Competitor / LG)에 따라 동작 방식이 다른 전체 세션 자동화 기능입니다.  
6개 PSM 모드를 순서대로 자동으로 돌며 각 시퀀스를 연속 실행합니다.

### 측정 순서 (공통)

| 순서 | PSM 모드 | 실행 시퀀스 |
|------|---------|-----------|
| 1 | SDR Vivid | 휘도 스윙 → APL 로딩 → **색재현율** → **명암비** |
| 2 | SDR Standard | 휘도 스윙 → APL 로딩 |
| 3 | SDR Cinema | 휘도 스윙 → APL 로딩 |
| 4 | HDR Vivid | 휘도 스윙 → APL 로딩 |
| 5 | HDR Standard | 휘도 스윙 → APL 로딩 |
| 6 | HDR Cinema | 휘도 스윙 → APL 로딩 |

> 색재현율·명암비는 SDR Vivid 1회만 측정합니다.

---

### Competitor 모드

모드 전환 시 **확인 다이얼로그**가 표시되며, 사용자가 TV PSM을 직접 변경 후 **[OK]** 를 누르면 다음 시퀀스가 자동 시작됩니다.

```
[다이얼로그] "SDR Vivid 로 변경 후 OK를 누르세요"
    → 휘도 스윙 자동 실행
    → APL 로딩 자동 실행
    → 색재현율 자동 실행   ← SDR Vivid 1회만
    → 명암비 자동 실행     ← SDR Vivid 1회만
[다이얼로그] "SDR Standard 로 변경 후 OK를 누르세요"
    → 휘도 스윙 자동 실행
    → APL 로딩 자동 실행
... (이하 동일)
```

- 각 단계마다 **[중지]** 버튼으로 전체 자동화 중단 가능
- 완료 시 `{brand}_{model}_all.xlsx` 자동 저장

---

### LG 모드

다이얼로그 없이 **시리얼 터미널 명령어**를 TV로 전송하여 PSM을 자동 전환합니다.  
측정 순서는 Competitor와 동일하며, 명령어 전송 코드 슬롯만 비워두고 나머지 자동화 흐름은 동일하게 동작합니다.

```python
def _send_psm_command(self, mode: str) -> None:
    """LG 전용 — PSM 변경 시리얼 명령 전송 슬롯.
    mode: "SDR_Vivid" | "SDR_Standard" | "SDR_Cinema" |
          "HDR_Vivid" | "HDR_Standard" | "HDR_Cinema"
    """
    # TODO: 여기에 시리얼 명령어를 채워 넣으세요
    # 예: self._lg_serial.write(b"...")
    pass
```

> 시리얼 포트 연결 설정은 [장비 연결] 패널에서 추가로 설정합니다.
