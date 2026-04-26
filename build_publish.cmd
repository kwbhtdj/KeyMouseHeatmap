@echo off
chcp 65001 >nul
setlocal

echo Building KeyMouseHeatmap...
dotnet restore
if errorlevel 1 goto fail

dotnet publish -c Release -r win-x64 --self-contained false -p:PublishSingleFile=false
if errorlevel 1 goto fail

echo.
echo Build succeeded.
echo EXE path:
echo %~dp0bin\Release\net9.0-windows\win-x64\publish\KeyMouseHeatmap.exe
goto end

:fail
echo.
echo Build failed. Please copy the red error text and send it to ChatGPT.

:end
pause
