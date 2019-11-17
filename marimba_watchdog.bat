''' >NUL  2>NUL
@echo off
cd /d %~dp0
:loop
python marimba_watchdog.py %*
goto loop
'''