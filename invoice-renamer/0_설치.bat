@echo off
cd /d "%~dp0"
echo 필요한 패키지를 설치합니다...
python -m pip install -r requirements.txt
echo.
echo 설치 끝. 이제 1_미리보기.bat 를 실행하세요.
pause
