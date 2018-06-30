@ECHO OFF
REM The following directory is for .NET 4.0
set DOTNETFX4=%SystemRoot%\Microsoft.NET\Framework\v4.0.30319
set PATH=%PATH%;%DOTNETFX4%
echo ---------------------------------------------------
regasm.exe "%~dp0..\src\UserNotifications\bin\Release\AddIn.UserNotification.dll" /codebase
echo ---------------------------------------------------
pause