@echo off
title Msg Generator Launcher

echo Running msg-generator...
echo.

java -jar msg-generator-1.0.0-shaded.jar

echo.
echo ==========================================
echo   Program finished.
echo   Press any key to exit...
echo ==========================================
echo.

pause > nul
