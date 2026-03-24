@echo off
title SBE Web Server
echo ====== Starting Shri Balaji Enterprise App ======
echo.
powershell.exe -ExecutionPolicy Bypass -File "%~dp0local_server.ps1"
exit
