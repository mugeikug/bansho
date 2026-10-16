@echo off
rem Add the installed Bansho app to the Explorer "Send to" menu.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0tools\register-sendto.ps1"
