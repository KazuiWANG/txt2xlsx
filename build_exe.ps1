# 打包 Windows 目录版 exe（输出到 dist\TXT重建Excel\）
# 用法：powershell -ExecutionPolicy Bypass -File .\build_exe.ps1
# 可用 -Python 指定解释器：.\build_exe.ps1 -Python "C:\Path\To\python.exe"
param(
    [string]$Python = ""
)
$ErrorActionPreference = "Stop"
$root = Split-Path -Parent $MyInvocation.MyCommand.Path
if (-not $Python) {
    $Python = $env:PYTHON
}
if (-not $Python) {
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd) { $Python = $cmd.Source }
}
if (-not $Python -or -not (Test-Path $Python)) {
    throw "未找到 Python 解释器。请设置环境变量 PYTHON，或用 -Python 指定 Python 3.9 的路径。"
}

Set-Location $root
& $Python -m PyInstaller --noconfirm --clean "$root\txt2xlsx.spec"
if ($LASTEXITCODE -ne 0) { throw "PyInstaller 构建失败，退出码 $LASTEXITCODE" }
Write-Host "构建完成：$root\dist\TXT重建Excel\TXT重建Excel.exe"
