@echo off
cd %~dp0

if "%1" == "" goto help

python %~dp0/script/dtxml2xlsx.py %1
goto end

:help
echo '���϶�dtxml�ļ�����������...'

:end
pause
