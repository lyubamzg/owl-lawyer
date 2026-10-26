# Dependency orchestration tests. No software is downloaded or installed.
[CmdletBinding()]
param([string]$Python = (Get-Command python -ErrorAction Stop).Source)
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('sova-setup-test-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $fixture | Out-Null
. (Join-Path $root 'packaging/windows/setup-dependencies.ps1') -InstallDir $fixture

function Assert-True { param($Value, [string]$Message); if (-not $Value) { throw $Message } }
Assert-True (Test-Python $Python) 'A working Python with Tkinter was not detected'
Assert-True (-not (Test-Python (Join-Path $fixture 'missing.exe'))) 'Missing Python accepted'
$result = Invoke-Checked $Python @('-c', "import sys; sys.stderr.write('diagnostic\n'); print('ok')") -Capture
Assert-True ($result -match 'diagnostic' -and $result -match 'ok') 'Successful native stderr was treated as failure'
$failed = $false
try { Invoke-Checked $Python @('-c', 'raise SystemExit(9)') } catch { $failed = $true }
Assert-True $failed 'Nonzero native exit code was ignored'

# Validate PowerShell 5.1-compatible parsing of every installer script.
foreach ($file in @(Get-ChildItem (Join-Path $root 'packaging/windows') -Filter '*.ps1')) {
    $tokens = $null; $parseErrors = $null
    $null = [Management.Automation.Language.Parser]::ParseFile($file.FullName, [ref]$tokens, [ref]$parseErrors)
    Assert-True ($parseErrors.Count -eq 0) "Parse failure: $($file.Name): $parseErrors"
}

$env:OS = 'Windows_NT' # Makes orchestration fixtures runnable on macOS too.
$script:packages = [Collections.Generic.List[string]]::new()
$script:commands = [Collections.Generic.List[string]]::new()
function Update-ProcessPath {}
function Install-WinGetPackage { param([string]$Id, [string]$Scope); $script:packages.Add($Id) }
function Find-Python { return $Python }
function Test-Python { param([string]$Path); return $true }
function Find-Tool { param([string]$Name); return (Join-Path $fixture "$Name.exe") }
function Install-RussianModel {}
function Invoke-Checked {
    param([string]$File, [string[]]$Arguments, [switch]$Capture)
    $script:commands.Add(($Arguments -join ' '))
    if ($Capture) { return "List of available languages (1):`nrus" }
}
function Get-DependencyStatus {
    return [ordered]@{
        python = $script:detected.python
        tesseract = $script:detected.tesseract
        poppler = $script:detected.poppler
        libreoffice = $script:detected.libreoffice
    }
}
function Reset-Case {
    $script:packages.Clear(); $script:commands.Clear()
    $script:detected = @{
        python = $Python; tesseract = (Join-Path $fixture 'tesseract.exe')
        poppler = (Join-Path $fixture 'pdftoppm.exe'); libreoffice = (Join-Path $fixture 'soffice.exe')
    }
    $script:InstallMissing = $false
    $script:InstallLibreOffice = $false
    $script:CheckOnly = $false
    $script:ReportFile = $null
}

Reset-Case
$script:detected.python = $null
$failed = $false
try { Initialize-SovaRuntime } catch { $failed = $_.Exception.Message -match 'Python' }
Assert-True $failed 'Missing Python without consent must stop setup'
Assert-True ($packages.Count -eq 0 -and $commands.Count -eq 0) 'Software installed without consent'

Reset-Case
Initialize-SovaRuntime
Assert-True ($packages.Count -eq 0) 'Existing components should be reused'
Assert-True (@($commands | Where-Object { $_ -match 'pip install.*--only-binary=:all:' }).Count -eq 1) 'Libraries must use the private runtime and binary wheels'
$config = Get-Content (Join-Path $fixture 'runtime-tools.json') -Raw | ConvertFrom-Json
Assert-True ($config.tool_dirs.Count -gt 0) 'Tool locations were not persisted'
Assert-True ($config.tessdata_dir -eq (Join-Path $fixture 'tessdata')) 'Wrong private OCR model location'

Reset-Case
$script:detected.libreoffice = $null
Initialize-SovaRuntime
Assert-True ($packages.Count -eq 0) 'Skipped LibreOffice must not be installed'

Reset-Case
$script:detected.libreoffice = $null
$script:InstallLibreOffice = $true
Initialize-SovaRuntime
Assert-True ($packages.Count -eq 1 -and $packages[0] -eq 'TheDocumentFoundation.LibreOffice') 'LibreOffice opt-in was not honored'

Reset-Case
$script:detected.python = $null
$script:detected.tesseract = $null
$script:detected.poppler = $null
$script:InstallMissing = $true
Initialize-SovaRuntime
Assert-True ($packages.Count -eq 3) 'Missing Python/OCR/PDF packages were not installed'
Assert-True ($packages -contains 'Python.Python.3.14') 'Wrong Python package'
Assert-True ($packages -contains 'UB-Mannheim.TesseractOCR') 'Wrong OCR package'
Assert-True ($packages -contains 'oschwartz10612.Poppler') 'Wrong PDF package'

Reset-Case
$script:CheckOnly = $true
Initialize-SovaRuntime
Assert-True ($packages.Count -eq 0 -and $commands.Count -eq 0) 'CheckOnly performed installations'

Reset-Case
function Invoke-Checked { param([string]$File, [string[]]$Arguments, [switch]$Capture); throw 'simulated install failure' }
$failed = $false
try { Initialize-SovaRuntime } catch { $failed = $_.Exception.Message -match 'simulated install failure' }
Assert-True $failed 'A failed dependency command must stop setup'
Write-Host 'PASS: Python detection, native exit codes/stderr, consent, reuse, LibreOffice opt-in, missing tools, read-only checks, and failure propagation'
