# 패널 측정 프로그램 (LED Measure)

CA-310 / CA-410 색채휘도계와 VG-876 / VG-879 패턴 제너레이터를 연동하여  
디스플레이 패널의 측정 시퀀스를 자동화하는 데스크탑 프로그램입니다.

---

## 목차

1. [프로젝트 구조](#프로젝트-구조)
2. [아키텍처](#아키텍처)
3. [측정 시퀀스](#측정-시퀀스)
4. [설치 및 실행](#설치-및-실행)
5. [EXE 빌드 (Windows)](#exe-빌드-windows)
6. [장비 사양](#장비-사양)
7. [데이터 기록 형식](#데이터-기록-형식)
8. [Excel 내보내기](#excel-내보내기)
9. [향후 개발 예정](#향후-개발-예정)

---

## 프로젝트 구조

```
led_measure/
├── core/                          ← UI 독립 공통 엔진
│   ├── engine.py                  MeasurementEngine — 시퀀스 오케스트레이터
│   ├── export.py                  Excel(.xlsx) 내보내기 (차트 포함)
│   ├── colorimetry.py             Robertson 보간법 기반 CCT / Duv 계산
│   ├── gamut_utils.py             u'v' 색역 면적 및 DCI/BT.2020 overlap 계산
│   ├── equipment/
│   │   ├── base.py                MeterBase, GeneratorBase ABC + 데이터클래스
│   │   ├── ca_meter.py            CA-310 / CA-410 pyserial 드라이버
│   │   ├── vg_generator.py        VG-876 / VG-879 패턴 제너레이터 드라이버
│   │   └── mock.py                하드웨어 없이 UI 테스트용 Mock 장비
│   └── sequences/
│       ├── center_align.py        1. 센터 맞추기
│       ├── lum_swing.py           2. 휘도 스윙 (연속 측정 스트림)
│       ├── lum_loading.py         3. APL 로딩
│       ├── gamut.py               4. 색재현율 (R→G→B→W→BK)
│       └── contrast.py            5. 명암비 (윈도우 크기별)
│
├── desktop/                       ← PySide6 데스크탑 UI
│   ├── main_window.py             MainWindow + 각 시퀀스 패널
│   └── worker.py                  QThread 래퍼 (progress·succeeded·error 시그널)
│
├── run_desktop.py                 ← 실행 진입점
├── led_measure.spec               PyInstaller 빌드 설정
├── build_windows.bat              Windows EXE 빌드 스크립트
└── requirements.txt
```

---

## 아키텍처

```
┌──────────────────────────────────────────┐
│              PySide6 Desktop             │
│  main_window.py   worker.py (QThread)    │
└────────────────────┬─────────────────────┘
                     │ import
┌────────────────────▼─────────────────────┐
│           Core Engine Layer              │
│  MeasurementEngine (engine.py)           │
│  ├── on_progress callback                │
│  ├── meter_lock (threading.Lock)         │
│  └── run_sequence(name, **kwargs)        │
│                                          │
│  Sequences            Equipment          │
│  ├── CenterAlign      ├── CaMeter        │
│  ├── LumSwing         ├── VgGenerator    │
│  ├── LumLoading       └── Mock(테스트용) │
│  ├── Gamut                               │
│  └── Contrast         Export (openpyxl)  │
└────────────────────┬─────────────────────┘
                     │ pyserial
┌────────────────────▼─────────────────────┐
│            Hardware Layer                │
│   CA-310 / CA-410     VG-876 / VG-879    │
└──────────────────────────────────────────┘
```

**핵심 원칙**: `core/` 는 `desktop/` 을 일절 import 하지 않습니다.  
UI가 바뀌어도 측정 로직은 그대로 재사용됩니다.

---

## 측정 시퀀스

### 1. 센터 맞추기 (Center Alignment)
- VG에서 흰색 1000×1000 px 센터 박스 출력 (3840×2160 화면 정중앙)
- 사용자가 측정기 렌즈를 화면 정중앙으로 조정 후 **[OK]** 클릭

### 2. 휘도 스윙 (Luminance Swing)
- 패턴: Black 바탕(BCOL4) + White 윈도우 (APL 10%, **H=31.6% × W=31.6%**)
- 연속 301회 측정 (1 sample/sec) → 시간 / Lv 실시간 라인 차트
- HDR 체크박스 토글 시 즉시 SHDMI4/SIF4/SHDR4 HDR10 출력 전환
- 저장 파일명: `lum_swing_[HDR/SDR]_[케이스]_[브랜드]_[모델].xlsx`

### 3. APL 로딩 (Luminance Loading)

| 버전 | 측정 포인트 수 |
|------|--------------|
| 37단계 | 1~100 구간 비선형 37포인트 |
| 10단계 | 1, 3, 10, 16, 25, 36, 49, 64, 81, 100 |
| 2단계  | 10, 100 |

- APL별 N회 연속 측정 (UI에서 1~20 설정) → 평균 또는 최대값 선택 표시
- 패턴: Black 바탕(BCOL4) + White 윈도우. 매 스텝 `_reset_window_memory()`로 라인 메모리 초기화
- 측정 항목: Lv, x, y, u', v', X, Y, Z, CCT, Duv
- APL vs Lv 인라인 차트 제공
- HDR 체크박스 토글 시 즉시 HDR10 출력 전환
- 저장 파일명: `lum_loading_[HDR/SDR]_[케이스]_[브랜드]_[모델].xlsx`

### 4. 색재현율 (Gamut)
- Full Pattern (100%) 순서: **R → G → B → W → BK** 각 1회 측정
- u'v' 색도 다이어그램: DCI-P3(파란 실선) / BT.2020(회색 실선) 기준 삼각형 표시
- 측정 삼각형(R-G-B) 및 각 색상 산점 표시
- DCI-P3 / BT.2020 overlap(%) 자동 계산 — Sutherland-Hodgman 클리핑 기반

### 5. 명암비 (Contrast Ratio)
- White 바탕(BCOL4 bg=255,255,255) + Black 윈도우
- 윈도우 크기(APL): **100% → 50% → 20% → 14.1% → 0%** 순 측정
  - 0% = 윈도우 없음 → 순수 White Raster (peak white 기준)
- HDR 체크박스 토글 시 즉시 HDR10 출력 전환

### 6. 보고서 템플릿
- 측정 완료된 모델 데이터를 누적하여 비교 표 생성
- 항목: White 휘도(HDR 10%/100%, SDR 10%/100%), Contrast Ratio, Black 수치, Color Gamut(DCI/BT.2020)
- 클립보드 복사 또는 Excel 내보내기

---

## 설치 및 실행

### 요구사항
- Python 3.10 이상
- 실제 측정 시 CA-310/410, VG-876/879 USB-Serial 연결

### 설치

```bash
python -m venv .venv

# Linux/macOS
source .venv/bin/activate

# Windows
.venv\Scripts\activate

pip install -r requirements.txt
```

### 실행

```bash
python run_desktop.py
```

> **Mock 모드**: 하드웨어 없이 테스트할 경우 상단 **[Mock 연결]** 클릭  
> 브랜드·모델명 미입력 시 Samsung / QN65S95D 로 자동 설정됩니다.

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

### 방법 3 — GitHub Actions

1. GitHub에 push
2. **Actions** 탭 → **Build Windows EXE** → **Run workflow**
3. 완료 후 **Artifacts** 에서 `led_measure-windows.zip` 다운로드

> PyInstaller 빌드는 플랫폼 종속적입니다 (Linux 빌드 → Linux 전용 / Windows 빌드 → `.exe`).

---

## 장비 사양

### CA-410 시리얼 통신 파라미터

| 파라미터 | 값 |
|---------|-----|
| Baud Rate | 38400 |
| Data Bits | 7 |
| Parity | EVEN |
| Stop Bits | 2 |
| Flow Control | RTS/CTS |

- 측정 명령: `MES\r\n`
- 응답 형식: `Lv, x, y, u', v', X, Y, Z, CCT, Duv` (10필드 CSV)
- CA-310은 8필드 응답 (CCT/Duv 없음 → 0.0 패딩)

### Mock 장비 시뮬레이션 타이밍

| 조건 | 대기 시간 |
|------|---------|
| 패턴 변경 후 디스플레이 안정화 | 2.0 s |
| CA-410 측정 (Lv ≥ 1 cd/m²) | 1.5 s |
| CA-410 측정 (0.01 ≤ Lv < 1) | 2.0 s |
| CA-410 측정 (Lv < 0.01) | 3.0 s |

### VG-876 / VG-879 serial protocol

바이너리 프레임: `STX(02H) + FDH + CMD1 + CMD2 + ASCII 파라미터(콤마 구분) + ETX(03H)`

| 커맨드 | 코드 | 설명 |
|--------|------|------|
| EXPDN4 | `24H 20H` | 프로그램 실행 / 타이밍 로드 |
| WINDOW4 | `28H 61H` | 윈도우 좌표 등록 (H/V 라인 메모리에 누적 저장) |
| WINCOL4 | `28H 62H` | 윈도우 내부 색상 (R/G/B: 0-65535, Bit Mode: 8-16) |
| WINCLR4 | `28H 63H` | H/V 라인 메모리 전체 클리어 (장비 정상 상태에서 안정적 — `_reset_window_memory()` 내부 사용) |
| BCOL4   | `28H 71H` | 윈도우 바깥 배경 색상 |
| SHDMI4  | `20H 36H` | HDMI 출력 포맷 (모드/색공간/비트심도) |
| SIF4    | `20H 38H` | AVI InfoFrame (BT.2020 색역) |
| SHDR4   | `20H C5H` | HDR10 Dynamic Range 메타데이터 |

#### Python API

- `show_full_field(r, g, b, bit_mode=8)` — `_reset_window_memory` → WINDOW4(전체) → WINCOL4 → EXPDN4
- `show_window_patch(width_pct, height_pct, r, g, b, bg_r=0, bg_g=0, bg_b=0, bit_mode=8)` — `_reset_window_memory` → BCOL4(배경) → WINDOW4(중앙) → WINCOL4 → EXPDN4
- `show_center_align()` — `_reset_window_memory` → WINDOW4(1000×1000 센터) → WINCOL4(흰색) → EXPDN4
- `show_crosshair()` — `_reset_window_memory` → BCOL4(black) → WINDOW4(수평) → WINDOW4(수직) → WINCOL4(white) → EXPDN4
- `set_hdr(enabled)` / `set_sdr()` — SHDMI4 + SIF4 + SHDR4 → EXPDN4
- `reset()` — ENQ 터미널 재진입 + EXPDN4(2286,0) 컬러바 복귀

> **_reset_window_memory**: WINCLR4 [28H 63H] 로 H/V 라인 메모리 전체 클리어.  
> 더미 1×1 창 방식(B방식)은 메모리를 실제로 클리어하지 않아 이전 APL 윈도우가 누적되는 문제가 있어 WINCLR4로 복귀.  
> `test_pattern_sequence.py` 테스트 1(APL 크기 변경), 테스트 2(gamut 색상 전환)로 검증 완료.

---

## 데이터 기록 형식

모든 시퀀스에서 `MeasureResult` 데이터클래스로 측정값을 통일합니다.

```python
@dataclass
class MeasureResult:
    timestamp_ms: int      # 측정 시각 (Unix ms)
    Lv:           float    # 휘도 cd/m²
    x:            float    # CIE x
    y:            float    # CIE y
    u_prime:      float    # CIE u'
    v_prime:      float    # CIE v'
    X:            float    # CIE X
    Y:            float    # CIE Y (= Lv)
    Z:            float    # CIE Z
    cct:          float    # 상관 색온도 K (장비 출력값)
    duv:          float    # Planckian locus 거리 (장비 출력값)
    pattern_info: PatternInfo
```

---

## Excel 내보내기

| 메서드 | 파일명 형식 | 내용 |
|--------|-----------|------|
| `export_lum_swing` | `lum_swing_[HDR/SDR]_[케이스]_[브랜드]_[모델].xlsx` | 케이스별 시트 + 시간-Lv 차트 |
| `export_lum_loading` | `lum_loading_[HDR/SDR]_[케이스]_[브랜드]_[모델].xlsx` | APL 요약 + 원시 데이터 시트 |
| `export_gamut` | `gamut_[브랜드]_[모델].xlsx` | R/G/B/W/BK 컬러별 측정값 + u'v' 색도 차트 |
| `export_contrast` | `contrast_[브랜드]_[모델].xlsx` | 윈도우 크기별 Lv + 명암비 |
| `export_report_template` | `report_[브랜드]_[모델].xlsx` | 다중 모델 비교 요약 표 |

- Excel 시트 내 `Time (s)` 컬럼: 측정 순번 (1, 2, 3 …)
- 브랜드별 그래프 색상: LG → 빨간색, Samsung → 파란색, 기타 → 초록색 등

---

## 향후 개발 예정

- [ ] CA-310 파싱 포맷 검증 (CA-410과 응답 형식 차이 확인)
- [ ] 쿨링 타임 UI 설정 (현재 하드코딩 3초)
- [ ] 측정 이력 DB 저장 (SQLite)
- [ ] 보고서 PDF 출력
- [ ] 아이콘 및 설치 인스톨러 (NSIS / Inno Setup)

---

## 변경 이력 (Changelog)

### 2026-05-28 (3)

- **[수정] 컬러바 배경 — ALLCLR4 도입** (`vg_generator.py`)
  - 원인: WINCLR4 는 윈도우 플레인만 클리어, 컬러바 베이스(program 2286) 잔존 → BCOL4 무효화
  - 수정: `_reset_window_memory()` → ALLCLR4 [28H 60H] (All planes clear)
  - 결과: ALLCLR4 후 BCOL4(black) 정상 작동 → 측정 패턴 배경 블랙 표시

- **[수정] gamut 깜빡임 및 속도 개선** (`vg_generator.py`)
  - `show_full_field()` 최적화: 처음만 ALLCLR4+WINDOW4(전체) 등록, 이후 WINCOL4만으로 즉시 색상 전환
  - 매뉴얼 확인: WINCOL4는 전역 색상이며 등록 창 전체에 즉시 적용

- **[수정] 명암비 측정 동작 안 함** (`engine.py`)
  - 원인: `_run_contrast()` → `_WINDOW_SIZES_PCT` import 오류 (실제 변수명: `_APL_SIZES_PCT`)
  - 수정: `from .sequences.contrast import _APL_SIZES_PCT`

- **[수정] 센터 맞추기 패턴 개선** (`vg_generator.py`)
  - 컬러바 베이스 유지(WINCLR4) + 중앙 박스(ㅁ) + 수평선 + 수직선 오버레이 → ㅁ+X 패턴
  - `show_crosshair()` 도 동일 방식: 컬러바 + 흰색 십자선

### 2026-05-28
- **[수정] 휘도 스윙 윈도우 크기 명시화** (`lum_swing.py`)
  - 패턴: 블랙 바탕 + 흰색 윈도우, W=31.6% / H=31.6% (10% APL) 고정값으로 지정
  - 기존: `sqrt(0.10) × 100 = 31.62%` 계산값 → 변경: `31.6%` 규격 고정

- **[수정] VG 윈도우 메모리 초기화 방식** (`vg_generator.py`)
  - 신규 내부 헬퍼 `_reset_window_memory()` 추가 — WINCLR4 [28H 63H] 로 H/V 라인 메모리 클리어
  - 전체 패턴 헬퍼(`show_full_field`, `show_window_patch`, `show_center_align`, `show_crosshair`)에 적용
  - `show_crosshair()` 미구현 → 구현 완료 (수평+수직 흰색 라인, 블랙 배경)
  - ※ 더미 1×1 창 방식(B방식)은 메모리를 클리어하지 않아 lum_loading APL 크기 미변경 문제 발생 → WINCLR4 방식으로 유지
  - ※ `test_fullfield_diag.py` A방식(WINCLR4) 실패는 테스트 시작 시 장비 freeze 상태 때문이었음 — WINCLR4 자체 문제 아님
  - `test_pattern_sequence.py` 테스트 1(APL 크기 변경) + 테스트 2(gamut 깜빡임) 검증 완료

- **[수정] gamut 색상 전환 시 컬러바 깜빡임 방지** (`vg_generator.py`)
  - 원인: 매 패턴 출력 전 `EXPDN4(2286,0)` 타이밍 재로드 → 컬러바 순간 표시
  - 수정: `_timing_loaded` 플래그 추가 — `set_sdr()` 또는 `reset()` 에서 1회 로드 후 이후 패턴에서 재로드 생략
  - `set_sdr()`: `_timing_loaded = False` → 강제 재로드 (HDR 해제 포함)
  - `disconnect()`: `_timing_loaded = False` 초기화

- **[추가] VG 장비 리셋 기능** (`vg_generator.py`, `main_window.py`)
  - `VgGenerator.reset()` — ENQ 터미널 재진입 + `EXPDN4(2286,0)` 컬러바 복귀 + 내부 상태 초기화
  - `GeneratorBase.reset()` — 기본 no-op 구현 (MockGenerator 등 하위 클래스는 오버라이드 불필요)
  - UI 연결 패널에 **[장비 리셋]** 주황색 버튼 추가 (VG 연결 시 활성화, 백그라운드 스레드 실행)

- **[수정] QThread 크래시 수정** (`worker.py`, `main_window.py`)
  - `finished` 시그널명 → `succeeded` 로 변경 (QThread::finished C++ 시그널 shadowing 방지)
  - 모든 워커에 `worker.finished.connect(worker.deleteLater)` 추가
  - `_on_hdr_toggled` 에 `isRunning()` 가드 + `@Slot(int)` 데코레이터 추가

- **[수정] SPTS4 사용 제거** (`vg_generator.py`)
  - `SPTS4 [20H 2AH]`: 컬러바(코드 15) 외 패턴에서 R/G/B 출력 비활성화 확인 (`test_vg_pattern.py`)
  - `_load_init_pattern()` 에서 SPTS4 완전 제거, `EXPDN4(2286, 0)` 단독 사용
3. hdr 체크시 Hdr 전환 안됨... 
4. swing 때도 10% apl 패턴 아니고 지금 Full white 로 뜸 .. 
패턴제너레이터 컨트롤이 전혀 원하는 방향으로 안되어 답답함 