@ECHO OFF

REM The following directory is for .NET 4.0
set DOTNETFX2=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319
set PATH=%PATH%;%DOTNETFX2%

echo Installing Email Job Service...
echo ---------------------------------------------------
C:\Windows\Microsoft.NET\Framework\v4.0.30319\InstallUtil "%~dp0PowerBIExcelService.exe"
echo ---------------------------------------------------
echo Setting Service Recovery Options...
sc failure "PowerBIExcelService" reset= 300 actions= restart/20000/restart/20000/actions=""/1000
pause
echo ---------------------------------------------------
echo Done.