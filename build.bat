@echo off
echo ========================================
echo Project Scheduler - Build Executable
echo ========================================
echo.

echo Installing/updating dependencies...
pip install -r requirements.txt

echo.
echo Building executable...
python build_executable.py

echo.
echo Build process completed!
pause