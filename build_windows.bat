@echo off
echo ====================================
echo  LED Measure - Windows EXE Build
echo ====================================

:: Python 설치 확인
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] Python이 설치되어 있지 않습니다.
    echo https://www.python.org/downloads/ 에서 Python 3.12 설치 후 다시 실행하세요.
    pause
    exit /b 1
)

echo [1/4] 가상환경 생성 중...
if exist .venv_win rmdir /s /q .venv_win
python -m venv .venv_win

echo [2/4] 패키지 설치 중... (시간이 걸릴 수 있습니다)
.venv_win\Scripts\pip install --upgrade pip -q
.venv_win\Scripts\pip install PySide6 openpyxl pyserial python-pptx matplotlib lxml pyinstaller -q
if errorlevel 1 (
    echo [ERROR] 패키지 설치 실패
    pause
    exit /b 1
)

echo [3/4] EXE 빌드 중...
.venv_win\Scripts\pyinstaller led_measure.spec --clean --noconfirm
if errorlevel 1 (
    echo [ERROR] 빌드 실패
    pause
    exit /b 1
)

echo.
echo ====================================
echo  빌드 완료!
echo  실행 파일: dist\led_measure\led_measure.exe
echo ====================================
explorer dist\led_measure
pause
