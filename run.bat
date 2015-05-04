@echo off
set osarch=%1
echo OS architecture is %osarch%
echo.

if "%osarch%"=="x86" (
  echo 32-bit cscript
  echo.
  call C:\Windows\system32\cscript.exe //NOLOGO reg.vbs x86
) else (
  echo 32-bit cscript
  echo.
  call C:\Windows\SysWOW64\cscript.exe //NOLOGO reg.vbs x64
  echo.

  echo 64-bit cscript
  echo.
  call C:\Windows\system32\cscript.exe //NOLOGO reg.vbs x64
)
