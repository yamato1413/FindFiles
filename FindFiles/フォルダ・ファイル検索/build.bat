@echo off

C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe -target:winexe -reference:Microsoft.VisualBasic.dll %~dp0\FindFiles.cs -win32icon:%~dp0\shell32_55_4.ico
if %ERRORLEVEL% equ 0 (
    .\FindFiles.exe
)
