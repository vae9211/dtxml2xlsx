@echo off
cd %~dp0

if "%1" == "" goto help

python %~dp0/script/xlsx2dtxml.py %1
goto end

:help
echo '���϶�xlsx�ļ�����������...'

:end
pause
