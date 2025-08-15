@echo off
title Markdown to DOCX Converter

echo ===============================================
echo    MARKDOWN TO DOCX CONVERTER
echo ===============================================
echo.
echo Choose your conversion method:
echo.
echo 1. Simple Text Interface (Recommended)
echo 2. Web Interface (Fancy Browser Interface)
echo 3. Setup/Install Dependencies
echo 4. Exit
echo.

set /p choice="Enter your choice (1-4): "

if "%choice%"=="1" goto simple
if "%choice%"=="2" goto web  
if "%choice%"=="3" goto setup
if "%choice%"=="4" goto end
goto invalid

:simple
echo.
echo Starting simple converter...
python convert_drag_drop.py
pause
goto end

:web
echo.
echo Starting web interface...
echo Open your browser to: http://localhost:5000
python markdown_converter_web.py
pause
goto end

:setup
echo.
echo Running setup...
python EASY_SETUP.py
pause
goto end

:invalid
echo.
echo Invalid choice! Please enter 1, 2, 3, or 4.
pause
goto start

:end
echo.
echo Goodbye!
pause
