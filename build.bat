@echo off
REM build.bat — builds a self-contained Windows EXE
REM Requirements: .NET 8 SDK  https://dotnet.microsoft.com/download
REM
REM Usage:  build.bat
REM Output: publish\win-x64\NedapStockExporter.exe

SET RID=win-x64

echo Building for runtime: %RID%

dotnet publish NedapStockExporter.csproj ^
  --configuration Release ^
  --runtime %RID% ^
  --self-contained true ^
  /p:PublishSingleFile=true ^
  /p:EnableCompressionInSingleFile=true ^
  --output publish\%RID%

IF %ERRORLEVEL% EQU 0 (
  echo.
  echo [OK] Build successful!
  echo Output: publish\%RID%\NedapStockExporter.exe
  dir publish\%RID%\NedapStockExporter.exe
) ELSE (
  echo [FAIL] Build failed.
)
