@echo off
cd /d %~dp0
call .venv\Scripts\python tile_img_on_xl.py %*
if %ERRORLEVEL% NEQ 0 pause
