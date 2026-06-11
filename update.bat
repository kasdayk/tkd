@echo off
title Update Dashboard TKD - Kota Yogyakarta
cd /d "%~dp0"
set PATH=%PATH%;C:\Program Files\Git\cmd
echo.
echo  Update Dashboard TKD Kota Yogyakarta
echo  ======================================
echo.
python update.py
