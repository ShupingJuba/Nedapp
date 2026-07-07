#!/usr/bin/env bash
# build.sh — builds a self-contained Windows EXE
# Requirements: .NET 8 SDK  (https://dotnet.microsoft.com/download)
#
# Usage:
#   chmod +x build.sh
#   ./build.sh                   # Windows x64 (default)
#   ./build.sh linux-x64         # Linux binary
#   ./build.sh osx-x64           # macOS binary
#
# The resulting EXE lands in: publish/<rid>/NedapStockExporter.exe

RID=${1:-win-x64}

echo "Building for runtime: $RID"

dotnet publish NedapStockExporter.csproj \
  --configuration Release \
  --runtime "$RID" \
  --self-contained true \
  /p:PublishSingleFile=true \
  /p:EnableCompressionInSingleFile=true \
  --output "publish/$RID"

if [ $? -eq 0 ]; then
  echo ""
  echo "✅  Build successful!"
  echo "Output: publish/$RID/NedapStockExporter${RID:0:3/win/.exe}"
  ls -lh "publish/$RID/"
else
  echo "❌  Build failed."
fi
