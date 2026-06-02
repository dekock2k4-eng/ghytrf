@echo off
title SheetMind — Excel Add-in Installer
echo.
echo  SheetMind Excel Add-in Installer
echo  ===================================
echo.

:: Use the Python that's already in PATH (same env as the rest of the project)
python create_addin.py

echo.
pause
