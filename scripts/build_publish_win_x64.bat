@echo off
setlocal
cd /d "%~dp0\..\src\PptTplAgent"
echo Publishing single-file exe (win-x64)...
dotnet publish -c Release -r win-x64 /p:PublishSingleFile=true /p:SelfContained=true
echo Done.
pause
