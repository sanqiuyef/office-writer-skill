$ErrorActionPreference = 'Stop'
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
$dotnet = $env:DOTNET_EXE
if (-not $dotnet) {
    $candidates = @(
        'C:\Program Files\dotnet\dotnet.exe',
        (Get-Command dotnet -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source)
    )
    $dotnet = $candidates | Where-Object { $_ -and (Test-Path $_) } | Select-Object -First 1
}
if (-not $dotnet) {
    throw 'dotnet SDK 8.0+ 未找到；请安装 .NET SDK 或设置 DOTNET_EXE 指向 dotnet.exe'
}
$dist = Join-Path $root 'dist'
# PDFtoImage 依赖 pdfium.dll / libSkiaSharp.dll；单文件发布时把原生库
# 打进 exe 并按需自解压，避免 dist 只剩 exe 导致 OCR/渲染找不到 DLL。
& $dotnet publish $root -c Release -r win-x64 --self-contained false `
    -p:PublishSingleFile=true -p:IncludeNativeLibrariesForSelfExtract=true -o $dist
if ($LASTEXITCODE -ne 0) {
    throw "dotnet publish failed: exit $LASTEXITCODE"
}
Write-Host "published: $(Join-Path $dist 'office-native.exe')"
