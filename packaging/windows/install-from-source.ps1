#Requires -Version 5.1
[CmdletBinding()]
param([string]$InstallDir = (Join-Path $env:LOCALAPPDATA 'Programs\OwlLawyer'))
$ErrorActionPreference = 'Stop'
$root = Split-Path (Split-Path $PSScriptRoot -Parent) -Parent
$installMissing = (Read-Host 'Install missing Python/OCR/PDF components if needed? [Y/n]') -notmatch '^(n|no)$'
$installOffice = (Read-Host 'Install LibreOffice if it is missing? [Y/n]') -notmatch '^(n|no)$'
$arguments = @('-NoProfile', '-ExecutionPolicy', 'Bypass', '-File',
               (Join-Path $PSScriptRoot 'setup-dependencies.ps1'), '-InstallDir', $InstallDir)
if ($installMissing) { $arguments += '-InstallMissing' }
if ($installOffice) { $arguments += '-InstallLibreOffice' }
& powershell.exe @arguments
if ($LASTEXITCODE -ne 0) { exit $LASTEXITCODE }

$app = Join-Path $InstallDir 'app'
New-Item -ItemType Directory -Path $app -Force | Out-Null
Copy-Item (Join-Path $root 'app\*.py') -Destination $app -Force
New-Item -ItemType Directory -Path (Join-Path $app 'icons') -Force | Out-Null
Copy-Item (Join-Path $root 'app\icons\*.png') -Destination (Join-Path $app 'icons') -Force
foreach ($name in @('input-workbook.xlsx', 'demo-housing.xlsx',
                    'demo-family-questionnaire.xlsx', 'demo-consumer-questionnaire.xlsx')) {
    if (-not (Test-Path (Join-Path $app $name))) {
        Copy-Item -LiteralPath (Join-Path $root "app\$name") -Destination $app
    }
}
foreach ($directory in @('templates', 'incoming')) {
    $source = Join-Path $root $directory
    foreach ($file in @(Get-ChildItem $source -Recurse -File)) {
        $relative = $file.FullName.Substring($source.Length).TrimStart('\')
        $target = Join-Path (Join-Path $InstallDir $directory) $relative
        New-Item -ItemType Directory -Path (Split-Path $target -Parent) -Force | Out-Null
        if (-not (Test-Path -LiteralPath $target)) { Copy-Item -LiteralPath $file.FullName -Destination $target }
    }
}
Copy-Item -LiteralPath (Join-Path $PSScriptRoot 'launch.cmd') -Destination (Join-Path $InstallDir 'OwlLawyer.cmd') -Force

# Wrap the original 256px PNG in an ICO file for a proper Windows shortcut.
$png = [IO.File]::ReadAllBytes((Join-Path $root 'icons\owl-lawyer-256.png'))
$iconPath = Join-Path $InstallDir 'owl-lawyer.ico'
$writer = [IO.BinaryWriter]::new([IO.File]::Create($iconPath))
try {
    $writer.Write([UInt16]0); $writer.Write([UInt16]1); $writer.Write([UInt16]1)
    1..4 | ForEach-Object { $writer.Write([byte]0) }
    $writer.Write([UInt16]1); $writer.Write([UInt16]32)
    $writer.Write([UInt32]$png.Length); $writer.Write([UInt32]22); $writer.Write($png)
} finally { $writer.Dispose() }
$shell = New-Object -ComObject WScript.Shell
foreach ($folder in @([Environment]::GetFolderPath('Desktop'),
                       [Environment]::GetFolderPath('Programs'))) {
    $shortcut = $shell.CreateShortcut((Join-Path $folder 'SOVA.lnk'))
    $shortcut.TargetPath = Join-Path $InstallDir 'runtime\Scripts\pythonw.exe'
    $shortcut.Arguments = '"' + (Join-Path $app 'desktop_app.py') + '"'
    $shortcut.WorkingDirectory = $app
    $shortcut.IconLocation = $iconPath
    $shortcut.Save()
}
Start-Process -FilePath (Join-Path $InstallDir 'runtime\Scripts\pythonw.exe') -ArgumentList ('"' + (Join-Path $app 'desktop_app.py') + '"') -WorkingDirectory $app
Write-Host "SOVA installed in $InstallDir"
