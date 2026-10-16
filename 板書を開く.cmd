@echo off
rem Open the board app in Microsoft Edge "app mode" (no address bar or tabs).
rem Works from any folder (uses this file's own location), e.g. after copying to a USB stick.
start "" /min powershell -NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "%~dp0tools\launch-edge.ps1" %*
