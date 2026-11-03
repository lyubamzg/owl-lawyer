# Direct-download regression tests. No external installers are executed.
[CmdletBinding()]
param()
$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$fixture = Join-Path ([IO.Path]::GetTempPath()) ('sova-download-test-' + [guid]::NewGuid())
New-Item -ItemType Directory -Path $fixture | Out-Null
. (Join-Path $root 'packaging/windows/setup-dependencies.ps1') -InstallDir $fixture
$directPackageFunction = ${function:Install-DirectPackage}
function Assert-True { param($Value, [string]$Message); if (-not $Value) { throw $Message } }
$env:LOCALAPPDATA = $fixture

foreach ($id in @('Python.Python.3.14', 'UB-Mannheim.TesseractOCR',
                  'oschwartz10612.Poppler', 'TheDocumentFoundation.LibreOffice',
                  'Microsoft.VCRedist.2015+.x64')) {
    $package = Get-DependencyPackage $id
    Assert-True ($package.url -like 'https://*' -and $package.sha256.Length -eq 64) "Bad package metadata: $id"
}

# Correct cached downloads are reused without connecting to the network.
$cache = Join-Path $fixture 'OwlLawyer/Downloads'
New-Item -ItemType Directory -Path $cache -Force | Out-Null
$cached = Join-Path $cache 'fixture.exe'
[IO.File]::WriteAllText($cached, 'verified fixture')
$package = [pscustomobject]@{
    filename = 'fixture.exe'; url = 'https://example.invalid/fixture.exe'
    sha256 = (Get-FileHash $cached -Algorithm SHA256).Hash.ToLower()
}
function Invoke-WebRequest { throw 'Unexpected network request' }
Assert-True ((Get-VerifiedDownload $package) -eq $cached) 'Verified cache was not reused'

# A bad download cannot reach an installer or replace a verified cache entry.
function Invoke-WebRequest { param($Uri, $OutFile, $TimeoutSec, [switch]$UseBasicParsing); [IO.File]::WriteAllText($OutFile, 'corrupt download') }
$package.sha256 = '0' * 64
$failed = $false
try { Get-VerifiedDownload $package } catch { $failed = $_.Exception.Message -match 'Checksum verification failed' }
Assert-True $failed 'A corrupt download was accepted'
Assert-True (@(Get-ChildItem $cache -Filter '*.partial').Count -eq 0) 'Failed partial downloads were left behind'

# Windows 1803 must never depend on WinGet, even if a stray executable exists.
$script:build = 17134
$script:hasWinget = $true
$script:backend = ''
function Get-WindowsBuild { return $script:build }
function Get-Command {
    param([string]$Name, $ErrorAction)
    if ($script:hasWinget) { return [pscustomobject]@{ Source = 'winget.exe' } }
    return $null
}
function Update-ProcessPath {}
function Install-DirectPackage { param([string]$Id); $script:backend = 'direct' }
function Install-WinGetPackage { param([string]$Id, [string]$Scope); $script:backend = 'winget' }
Install-Dependency 'Python.Python.3.14' 'user'
Assert-True ($backend -eq 'direct') 'Windows 1803 incorrectly tried WinGet'
$script:build = 17763
Install-Dependency 'Python.Python.3.14' 'user'
Assert-True ($backend -eq 'winget') 'Modern Windows should reuse available WinGet'
$script:hasWinget = $false
Install-Dependency 'Python.Python.3.14' 'user'
Assert-True ($backend -eq 'direct') 'Missing WinGet must use direct downloads'

# Successful, restart-required, and rejected elevated installer exits.
$script:installerExit = 0
function Start-Process {
    param($FilePath, $ArgumentList, $Verb, [switch]$Wait, [switch]$PassThru)
    $script:lastInstallerArguments = $ArgumentList
    $script:lastInstallerFile = $FilePath
    $script:lastInstallerVerb = $Verb
    return [pscustomobject]@{ ExitCode = $script:installerExit }
}
Invoke-PackageInstaller 'fixture.exe' '/passive' -Elevate
$script:installerExit = 3010
Invoke-PackageInstaller 'fixture.exe' '/passive'
Assert-True (Test-Path (Join-Path $fixture 'setup-restart-required')) 'Restart request was lost'
$script:installerExit = 1602
$failed = $false
try { Invoke-PackageInstaller 'fixture.exe' '/passive' -Elevate } catch { $failed = $_.Exception.Message -match 'exit 1602' }
Assert-True $failed 'Cancelled installer was treated as successful'

# Exercise actual direct-package routing, including private Poppler extraction.
Set-Item function:Install-DirectPackage -Value $directPackageFunction
$script:installerExit = 0
$env:SystemRoot = $fixture
function Test-VCRuntime { param([string]$MinimumVersion); return $true }
$staging = Join-Path $fixture 'archive/poppler-25.07.0/Library/bin'
New-Item -ItemType Directory -Path $staging -Force | Out-Null
[IO.File]::WriteAllText((Join-Path $staging 'pdftoppm.exe'), 'test binary')
$script:popplerArchive = Join-Path $fixture 'poppler.zip'
Compress-Archive -Path (Join-Path $fixture 'archive/*') -DestinationPath $popplerArchive
function Get-VerifiedDownload {
    param($Package)
    if ($Package.kind -eq 'poppler') { return $script:popplerArchive }
    return $cached
}
Install-DirectPackage 'Python.Python.3.14'
Assert-True ($lastInstallerArguments -match 'Include_tcltk=1' -and
             $lastInstallerArguments -match 'InstallAllUsers=0' -and
             -not $lastInstallerVerb) 'Python must install per-user with Tcl/Tk'
Install-DirectPackage 'UB-Mannheim.TesseractOCR'
Assert-True ($lastInstallerArguments -eq '/S' -and $lastInstallerVerb -eq 'RunAs') 'Tesseract installation flags are incorrect'
Install-DirectPackage 'TheDocumentFoundation.LibreOffice'
Assert-True ($lastInstallerFile -like '*msiexec.exe' -and
             $lastInstallerArguments -match '/passive /norestart' -and
             $lastInstallerVerb -eq 'RunAs') 'LibreOffice must use the system MSI installer'
Install-DirectPackage 'oschwartz10612.Poppler'
Assert-True (Test-Path (Join-Path $fixture 'tools/poppler/poppler-25.07.0/Library/bin/pdftoppm.exe')) 'Poppler was not extracted into its private tool folder'
Write-Host 'PASS: verified cache, checksum rejection, Windows 1803 without WinGet, modern fallback, native installer exits and restarts'
exit 0
