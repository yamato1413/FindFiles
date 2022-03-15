@echo off

C:\Windows\Microsoft.NET\Framework64\v4.0.30319\csc.exe -target:winexe %~dp0\FindFiles.cs 
if %ERRORLEVEL% equ 0 (
    .\FindFiles.exe
)
