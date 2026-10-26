#Requires -Version 5.1
[CmdletBinding()]
param([string]$Version = '0.2.0')
$ErrorActionPreference = 'Stop'
if ($Version -notmatch '^\d+\.\d+\.\d+(?:\.\d+)?$') { throw 'Use a numeric version, e.g. 0.2.0' }
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
& python (Join-Path $PSScriptRoot 'build_assets.py')
if ($LASTEXITCODE -ne 0) { throw 'Asset preparation failed' }
$compiler = Get-Command ISCC.exe -ErrorAction SilentlyContinue
if ($compiler) {
    $iscc = $compiler.Source
} else {
    $iscc = Join-Path ${env:ProgramFiles(x86)} 'Inno Setup 6\ISCC.exe'
    if (-not (Test-Path $iscc)) {
        $iscc = Join-Path $env:ProgramFiles 'Inno Setup 7\ISCC.exe'
    }
}
if (-not (Test-Path $iscc)) { throw 'Install Inno Setup 6.3+ (or add ISCC.exe to PATH).' }
& $iscc "/DAppVersion=$Version" (Join-Path $PSScriptRoot 'installer.iss')
if ($LASTEXITCODE -ne 0) { throw 'Inno Setup compilation failed' }
$artifact = Join-Path $root "dist\SOVA-Setup-$Version-x64.exe"
(Get-FileHash -LiteralPath $artifact -Algorithm SHA256).Hash.ToLower() + '  ' +
    (Split-Path $artifact -Leaf) | Set-Content -LiteralPath "$artifact.sha256" -Encoding ASCII
Write-Host "Built $artifact"
