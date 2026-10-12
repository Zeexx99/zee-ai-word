@echo off
title Zee AI for Word diagnostics
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0diagnose.ps1"
echo Diagnostics written to diag-log.txt - send it to Claude if the add-in does not load.
timeout /t 5
