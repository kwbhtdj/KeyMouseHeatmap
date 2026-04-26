@echo off
chcp 65001 >nul
setlocal

set OUT=%~dp0publish

echo Building KeyMouseHeatmap lightweight packages...
echo.
echo This version only creates 2 packages to avoid wasting disk space.
echo The lite package is framework-dependent and does not use single-file compression, which avoids NETSDK1176.
echo.

dotnet restore
if errorlevel 1 goto fail

if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"

echo.
echo [1/2] Lite x64 - smallest size, requires .NET 9 Desktop Runtime
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=false -p:PublishReadyToRun=false -p:DebugType=none -p:DebugSymbols=false -p:EnableCompressionInSingleFile=false -o "%OUT%\lite-win-x64"
if errorlevel 1 goto fail

echo.
echo [2/2] Portable compressed x64 - no .NET install required, single-file compressed package
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:PublishReadyToRun=false -p:DebugType=none -p:DebugSymbols=false -p:EnableCompressionInSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o "%OUT%\portable-compressed-win-x64"
if errorlevel 1 goto fail

echo.
echo Build succeeded.
echo.
echo Smallest package:
echo %OUT%\lite-win-x64

echo.
echo No-runtime portable package:
echo %OUT%\portable-compressed-win-x64\KeyMouseHeatmap.exe

echo.
echo Tip: If you already installed .NET 9 Desktop Runtime, release lite-win-x64.
echo If users do not want to install .NET, release portable-compressed-win-x64.
goto end

:fail
echo.
echo Build failed. Please copy the red error text and send it to ChatGPT.

:end
pause
