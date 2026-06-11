@echo off
title Update Manual Dashboard TKD - Kota Yogyakarta
cd /d "%~dp0"
set PATH=%PATH%;C:\Program Files\Git\cmd
echo.
echo  ================================================
echo   Update Manual Dashboard TKD Kota Yogyakarta
echo  ================================================
echo.
echo  Pastikan file XLS dari SIKD sudah ada di folder ini!
echo.
python update_manual.py
