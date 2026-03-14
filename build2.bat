@echo off
setlocal

set MSVC=C:\Program Files\Microsoft Visual Studio\2022\Community\VC\Tools\MSVC\14.43.34808
set SDK=C:\Program Files (x86)\Windows Kits\10
set SDKVER=10.0.22621.0

set INCLUDE=%MSVC%\include;%SDK%\Include\%SDKVER%\ucrt;%SDK%\Include\%SDKVER%\shared;%SDK%\Include\%SDKVER%\um
set LIB=%MSVC%\lib\x64;%SDK%\Lib\%SDKVER%\ucrt\x64;%SDK%\Lib\%SDKVER%\um\x64

"%MSVC%\bin\Hostx64\x64\cl.exe" SimpleBrowser.cpp /DUNICODE /D_UNICODE /EHsc /W3 ^
  /link ole32.lib oleaut32.lib uuid.lib user32.lib gdi32.lib

echo.
echo EXIT_CODE=%ERRORLEVEL%
endlocal
