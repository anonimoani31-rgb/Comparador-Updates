@echo off
cd /d "%~dp0"
"%~dp0jdk\bin\java.exe" -cp ".;lib/*" Main
pause