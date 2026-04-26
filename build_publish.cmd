@echo off
chcp 65001 >nul
setlocal

echo Building KeyMouseHeatmap packages...
dotnet restore
if errorlevel 1 goto fail

set OUT=%~dp0publish
if exist "%OUT%" rmdir /s /q "%OUT%"
mkdir "%OUT%"

echo.
echo [1/4] Framework-dependent x64 (requires .NET 9 Desktop Runtime)
dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=false -o "%OUT%\framework-dependent-win-x64"
if errorlevel 1 goto fail

echo.
echo [2/4] Self-contained x64 folder (no .NET install required)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -o "%OUT%\self-contained-win-x64"
if errorlevel 1 goto fail

echo.
echo [3/4] Self-contained x64 single exe (no .NET install required)
dotnet publish -c Release -r win-x64 --self-contained true -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o "%OUT%\singlefile-win-x64"
if errorlevel 1 goto fail

echo.
echo [4/4] Self-contained x86 folder (older 32-bit Windows)
dotnet publish -c Release -r win-x86 --self-contained true -p:PublishSingleFile=false -o "%OUT%\self-contained-win-x86"
if errorlevel 1 goto fail

echo.
echo Build succeeded.
echo Packages are in:
echo %OUT%
echo.
echo Recommended for most users without .NET:
echo %OUT%\singlefile-win-x64\KeyMouseHeatmap.exe
goto end

:fail
echo.
echo Build failed. Please copy the red error text and send it to ChatGPT.

:end
pause
