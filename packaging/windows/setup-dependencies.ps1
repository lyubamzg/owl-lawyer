#Requires -Version 5.1
[CmdletBinding()]
param(
    [Parameter(Mandatory = $true)][string]$InstallDir,
    [string]$RequirementsFile = (Join-Path $PSScriptRoot 'requirements.txt'),
    [switch]$InstallMissing,
    [switch]$InstallLibreOffice,
    [switch]$CheckOnly,
    [string]$RussianModel = (Join-Path $PSScriptRoot 'rus.traineddata'),
    [string]$DependencyManifest = (Join-Path $PSScriptRoot 'dependency-manifest.json'),
    [string]$ReportFile
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Update-ProcessPath {
    $paths = @(
        [Environment]::GetEnvironmentVariable('PATH', 'Machine'),
        [Environment]::GetEnvironmentVariable('PATH', 'User'),
        $env:PATH,
        (Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Links')
    )
    $config = Join-Path $InstallDir 'runtime-tools.json'
    if (Test-Path -LiteralPath $config) {
        try {
            $saved = Get-Content -LiteralPath $config -Raw | ConvertFrom-Json
            $paths += @($saved.tool_dirs)
        } catch { Write-Host 'Ignoring an invalid previous tool configuration; setup will rebuild it.' }
    }
    $directories = foreach ($value in @($paths | Where-Object { $_ })) {
        foreach ($directory in ($value -split ';')) {
            if ($directory.Trim()) { [Environment]::ExpandEnvironmentVariables($directory.Trim()) }
        }
    }
    $env:PATH = (@($directories | Select-Object -Unique) -join ';')
}

function Invoke-Checked {
    param([string]$File, [string[]]$Arguments, [switch]$Capture)
    $previous = $ErrorActionPreference
    try {
        # PS 5.1 treats native stderr as ErrorRecord even on a successful exit.
        $ErrorActionPreference = 'Continue'
        $output = & $File @Arguments 2>&1
        $code = $LASTEXITCODE
    } finally { $ErrorActionPreference = $previous }
    if ($code -ne 0) {
        throw "Command failed ($code): $File $($Arguments -join ' ')`n$($output -join "`n")"
    }
    if ($Capture) { return ($output -join "`n") }
    $output | ForEach-Object { Write-Host $_ }
}

function Test-Python {
    param([string]$Path)
    if (-not $Path -or -not (Test-Path -LiteralPath $Path)) { return $false }
    try {
        $probe = "import sys,struct,tkinter,venv,pip; assert (3,11)<=sys.version_info[:2]<(3,15); assert struct.calcsize('P')==8; print('SOVA_PYTHON_OK')"
        $result = & $Path -c $probe 2>$null
        return ($LASTEXITCODE -eq 0 -and $result -contains 'SOVA_PYTHON_OK')
    } catch { return $false }
}

function Find-Python {
    $candidates = @((Join-Path $InstallDir 'runtime\Scripts\python.exe'))
    foreach ($registry in @('HKCU:\Software\Python\PythonCore',
                            'HKLM:\Software\Python\PythonCore')) {
        if (Test-Path $registry) {
            foreach ($version in @(Get-ChildItem $registry | Sort-Object PSChildName -Descending)) {
                $key = Join-Path $version.PSPath 'InstallPath'
                if (Test-Path $key) {
                    $item = Get-Item $key
                    $path = $item.GetValue('ExecutablePath')
                    if (-not $path -and $item.GetValue('')) {
                        $path = Join-Path $item.GetValue('') 'python.exe'
                    }
                    if ($path) { $candidates += $path }
                }
            }
        }
    }
    foreach ($command in @('python.exe', 'python3.exe')) {
        $found = Get-Command $command -ErrorAction SilentlyContinue
        # Skip Windows Store aliases: invoking them can open the Store.
        if ($found -and $found.Source -notlike '*\WindowsApps\*') {
            $candidates += $found.Source
        }
    }
    foreach ($directory in @((Join-Path $env:LOCALAPPDATA 'Programs\Python'),
                              (Join-Path $env:LOCALAPPDATA 'Python'))) {
        if (Test-Path $directory) {
            $candidates += @(Get-ChildItem $directory -Filter python.exe -Recurse -File |
                             Select-Object -ExpandProperty FullName)
        }
    }
    foreach ($candidate in @($candidates | Select-Object -Unique)) {
        if (Test-Python $candidate) { return $candidate }
    }
    return $null
}

function Find-Tool {
    param([ValidateSet('tesseract', 'pdftoppm', 'soffice')][string]$Name)
    $found = Get-Command "$Name.exe" -ErrorAction SilentlyContinue
    if ($found) { return $found.Source }
    if ($Name -eq 'pdftoppm') {
        $private = Join-Path $InstallDir 'tools\poppler'
        if (Test-Path -LiteralPath $private) {
            $binary = Get-ChildItem $private -Filter pdftoppm.exe -Recurse -File |
                      Select-Object -First 1 -ExpandProperty FullName
            if ($binary) { return $binary }
        }
    }
    # Custom installation locations are often registered without a PATH entry.
    foreach ($hive in @('HKCU:', 'HKLM:')) {
        $keys = @("$hive\Software\Microsoft\Windows\CurrentVersion\App Paths\$Name.exe")
        if ($Name -eq 'soffice') { $keys += "$hive\Software\LibreOffice\UNO\InstallPath" }
        if ($Name -eq 'tesseract') { $keys += "$hive\Software\Tesseract-OCR" }
        foreach ($key in $keys) {
            if (Test-Path $key -ErrorAction SilentlyContinue) {
                $item = Get-Item $key
                $path = $item.GetValue('')
                if ($Name -eq 'tesseract' -and $key -like '*\Tesseract-OCR') {
                    $path = $item.GetValue('InstallDir')
                }
                if ($path) {
                    $path = $path.Trim('"')
                    if (Test-Path -LiteralPath $path -PathType Container) { $path = Join-Path $path "$Name.exe" }
                    if (Test-Path -LiteralPath $path -PathType Leaf) { return $path }
                }
            }
        }
    }
    $folders = @($env:ProgramFiles, ${env:ProgramFiles(x86)},
                 (Join-Path $env:LOCALAPPDATA 'Programs'))
    foreach ($base in @($folders | Where-Object { $_ })) {
        $relative = switch ($Name) {
            'tesseract' { 'Tesseract-OCR\tesseract.exe' }
            'soffice' { 'LibreOffice\program\soffice.exe' }
            'pdftoppm' { 'poppler\Library\bin\pdftoppm.exe' }
        }
        $candidate = Join-Path $base $relative
        if (Test-Path -LiteralPath $candidate) { return $candidate }
    }
    # WinGet portable packages are not always added to this process's PATH yet.
    if ($Name -eq 'pdftoppm') {
        foreach ($base in @((Join-Path $env:LOCALAPPDATA 'Microsoft\WinGet\Packages'),
                            (Join-Path $env:ProgramFiles 'WinGet\Packages'))) {
            if (Test-Path $base) {
                foreach ($package in @(Get-ChildItem $base -Directory -Filter 'oschwartz10612.Poppler*')) {
                    $candidate = Get-ChildItem $package.FullName -Recurse -File -Filter pdftoppm.exe |
                                 Select-Object -First 1 -ExpandProperty FullName
                    if ($candidate) { return $candidate }
                }
            }
        }
    }
    return $null
}

function Install-WinGetPackage {
    param([string]$Id, [string]$Scope)
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if (-not $winget) {
        throw 'Windows App Installer (WinGet) is missing. Install/update App Installer from https://aka.ms/getwinget, then run setup again. Alternatively install the missing tools manually.'
    }
    Write-Host "Installing $Id. Windows may request administrator approval."
    $arguments = @('install', '--id', $Id, '--exact', '--source', 'winget',
                   '--accept-package-agreements', '--accept-source-agreements',
                   '--silent', '--disable-interactivity', '--no-upgrade')
    if ($Scope) { $arguments += @('--scope', $Scope) }
    Invoke-Checked $winget.Source $arguments
    Update-ProcessPath
}

function Get-WindowsBuild {
    return [Environment]::OSVersion.Version.Build
}

function Get-DependencyPackage {
    param([string]$Id)
    $manifest = Get-Content -LiteralPath $DependencyManifest -Raw | ConvertFrom-Json
    $property = $manifest.PSObject.Properties[$Id]
    if (-not $property) { throw "No verified installer is configured for $Id." }
    $package = $property.Value
    if ($package.url -notmatch '^https://' -or $package.sha256 -notmatch '^[a-fA-F0-9]{64}$') {
        throw "Invalid download metadata for $Id."
    }
    return $package
}

function Get-VerifiedDownload {
    param($Package)
    $cache = Join-Path $env:LOCALAPPDATA 'OwlLawyer\Downloads'
    New-Item -ItemType Directory -Path $cache -Force | Out-Null
    $file = Join-Path $cache $Package.filename
    $expected = $Package.sha256.ToLower()
    if ((Test-Path -LiteralPath $file) -and
        (Get-FileHash -LiteralPath $file -Algorithm SHA256).Hash.ToLower() -eq $expected) {
        return $file
    }
    $partial = $file + '.' + [guid]::NewGuid() + '.partial'
    try {
        Write-Host "Downloading $($Package.filename)..."
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        $ProgressPreference = 'SilentlyContinue'
        Invoke-WebRequest -UseBasicParsing -Uri $Package.url -OutFile $partial -TimeoutSec 1800
        if ((Get-FileHash -LiteralPath $partial -Algorithm SHA256).Hash.ToLower() -ne $expected) {
            throw "Checksum verification failed for $($Package.filename). Nothing was installed from this download."
        }
        Move-Item -LiteralPath $partial -Destination $file -Force
        return $file
    } finally {
        if (Test-Path -LiteralPath $partial) { Remove-Item -LiteralPath $partial -Force }
    }
}

function Invoke-PackageInstaller {
    param([string]$File, [string]$Arguments, [switch]$Elevate)
    $options = @{ FilePath = $File; ArgumentList = $Arguments; Wait = $true; PassThru = $true }
    if ($Elevate) { $options.Verb = 'RunAs' }
    try { $process = Start-Process @options }
    catch { throw "Could not start the component installer. Allow the Windows administrator prompt and retry. $($_.Exception.Message)" }
    if ($process.ExitCode -eq 3010) {
        [IO.File]::WriteAllText((Join-Path $InstallDir 'setup-restart-required'), '1')
        Write-Host 'A component requested a Windows restart after setup.'
    } elseif ($process.ExitCode -ne 0) {
        throw "Component installation failed (exit $($process.ExitCode)): $File. See setup.log."
    }
}

function Test-VCRuntime {
    param([string]$MinimumVersion)
    $key = 'HKLM:\SOFTWARE\Microsoft\VisualStudio\14.0\VC\Runtimes\x64'
    if (Test-Path $key -ErrorAction SilentlyContinue) {
        $item = Get-ItemProperty $key
        if ($item.Installed -eq 1 -and [version]($item.Version.TrimStart('v')) -ge [version]$MinimumVersion) {
            return $true
        }
    }
    return $false
}

function Install-DirectPackage {
    param([string]$Id)
    $package = Get-DependencyPackage $Id
    if ($package.kind -eq 'vcredist' -and (Test-VCRuntime $package.version)) { return }
    if ($package.kind -in @('tesseract', 'poppler', 'libreoffice')) {
        Install-DirectPackage 'Microsoft.VCRedist.2015+.x64'
    }
    $file = Get-VerifiedDownload $package
    switch ($package.kind) {
        'python' {
            Invoke-PackageInstaller $file '/passive /norestart InstallAllUsers=0 PrependPath=0 Include_launcher=0 Include_test=0 Include_pip=1 Include_tcltk=1'
        }
        'tesseract' { Invoke-PackageInstaller $file '/S' -Elevate }
        'vcredist' { Invoke-PackageInstaller $file '/install /passive /norestart' -Elevate }
        'libreoffice' {
            $msiexec = Join-Path $env:SystemRoot 'System32\msiexec.exe'
            Invoke-PackageInstaller $msiexec ('/i "' + $file + '" /passive /norestart') -Elevate
        }
        'poppler' {
            $destination = Join-Path $InstallDir 'tools\poppler'
            New-Item -ItemType Directory -Path $destination -Force | Out-Null
            Expand-Archive -LiteralPath $file -DestinationPath $destination -Force
            if (-not (Get-ChildItem $destination -Filter pdftoppm.exe -Recurse -File | Select-Object -First 1)) {
                throw 'The verified Poppler archive did not contain pdftoppm.exe.'
            }
        }
        default { throw "Unsupported component installer type: $($package.kind)" }
    }
}

function Install-Dependency {
    param([string]$Id, [string]$Scope)
    $winget = Get-Command winget.exe -ErrorAction SilentlyContinue
    if ((Get-WindowsBuild) -ge 17763 -and $winget) {
        Install-WinGetPackage $Id $Scope
    } else {
        Write-Host 'Using verified direct downloads (WinGet is unavailable on this system).'
        Install-DirectPackage $Id
        Update-ProcessPath
    }
}

function Get-DependencyStatus {
    return [ordered]@{
        python = Find-Python
        tesseract = Find-Tool 'tesseract'
        poppler = Find-Tool 'pdftoppm'
        libreoffice = Find-Tool 'soffice'
    }
}

function Install-RussianModel {
    $directory = Join-Path $InstallDir 'tessdata'
    New-Item -ItemType Directory -Path $directory -Force | Out-Null
    $destination = Join-Path $directory 'rus.traineddata'
    $expected = 'e16e5e036cce1d9ec2b00063cf8b54472625b9e14d893a169e2b0dedeb4df225'
    if (-not (Test-Path -LiteralPath $RussianModel)) {
        $RussianModel = Join-Path ([IO.Path]::GetTempPath()) ('sova-rus-' + [guid]::NewGuid() + '.traineddata')
        [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
        Invoke-WebRequest -UseBasicParsing -Uri 'https://raw.githubusercontent.com/tesseract-ocr/tessdata_fast/4.1.0/rus.traineddata' -OutFile $RussianModel
    }
    if ((Get-FileHash -LiteralPath $RussianModel -Algorithm SHA256).Hash.ToLower() -ne $expected) {
        throw 'Russian OCR model verification failed. Download the installer again.'
    }
    Copy-Item -LiteralPath $RussianModel -Destination $destination -Force
    $env:TESSDATA_PREFIX = $directory
}

function Initialize-SovaRuntime {
    if (-not [Environment]::Is64BitOperatingSystem -or $env:OS -ne 'Windows_NT' -or (Get-WindowsBuild) -lt 17134) {
        throw 'SOVA setup requires 64-bit Windows 10 version 1803 (build 17134) or later.'
    }
    Update-ProcessPath
    $status = Get-DependencyStatus
    if ($CheckOnly) {
        $json = $status | ConvertTo-Json
        if ($ReportFile) { [IO.File]::WriteAllText($ReportFile, $json, [Text.UTF8Encoding]::new($false)) }
        Write-Host $json
        return
    }

    if (-not $status.python) {
        if (-not $InstallMissing) { throw 'Python 3.11-3.14 (64-bit), including Tkinter and pip, is required. Enable installation of missing components or install Python manually.' }
        Install-Dependency 'Python.Python.3.14' 'user'
        $status.python = Find-Python
        if (-not $status.python) { throw 'Python was installed but its Tkinter/pip runtime could not be detected. Repair Python with the Tcl/Tk and pip options enabled, then retry setup.' }
    }
    foreach ($tool in @(
        @{ Key = 'tesseract'; Name = 'tesseract'; Package = 'UB-Mannheim.TesseractOCR'; Scope = '' },
        @{ Key = 'poppler'; Name = 'pdftoppm'; Package = 'oschwartz10612.Poppler'; Scope = 'user' }
    )) {
        if (-not $status[$tool.Key]) {
            if (-not $InstallMissing) { throw "$($tool.Name) is missing. Enable installation of missing components or install it manually." }
            Install-Dependency $tool.Package $tool.Scope
            $status[$tool.Key] = Find-Tool $tool.Name
            if (-not $status[$tool.Key]) { throw "$($tool.Name) is still unavailable after installation. See setup.log." }
        }
    }
    if (-not $status.libreoffice -and $InstallLibreOffice) {
        Install-Dependency 'TheDocumentFoundation.LibreOffice' ''
        $status.libreoffice = Find-Tool 'soffice'
        if (-not $status.libreoffice) { throw 'LibreOffice was not detected after installation. Retry setup after a reboot if Windows requested one.' }
    }

    $runtime = Join-Path $InstallDir 'runtime'
    $python = Join-Path $runtime 'Scripts\python.exe'
    if (-not (Test-Python $python)) {
        Write-Host 'Creating the private SOVA Python environment...'
        Invoke-Checked $status.python @('-m', 'venv', $runtime)
    }
    Write-Host 'Checking/installing SOVA Python libraries...'
    Invoke-Checked $python @('-m', 'pip', 'install', '--disable-pip-version-check',
                              '--only-binary=:all:', '--timeout', '60', '--retries', '2',
                              '-r', $RequirementsFile)
    Invoke-Checked $python @('-c', "import tkinter,openpyxl,docx,lxml; print('SOVA libraries: OK')")
    Invoke-Checked $status.tesseract @('--version')
    Invoke-Checked $status.poppler @('-v')
    Install-RussianModel
    $languages = Invoke-Checked $status.tesseract @('--list-langs') -Capture
    if ($languages -notmatch '(?m)^rus\s*$') { throw 'Tesseract could not load the Russian OCR model.' }

    $dirs = @($status.tesseract, $status.poppler, $status.libreoffice) |
            Where-Object { $_ } | ForEach-Object { Split-Path $_ -Parent } | Select-Object -Unique
    $config = [ordered]@{
        tool_dirs = @($dirs)
        tessdata_dir = (Join-Path $InstallDir 'tessdata')
        soffice = $status.libreoffice
    }
    [IO.File]::WriteAllText((Join-Path $InstallDir 'runtime-tools.json'),
        ($config | ConvertTo-Json -Depth 3), [Text.UTF8Encoding]::new($false))
    if (-not $status.libreoffice) {
        Write-Host 'LibreOffice was skipped. Install it later to enable editing/conversion, then rerun setup.'
    }
    Write-Host 'SOVA dependencies are ready.'
}

if ($MyInvocation.InvocationName -ne '.') {
    $logDirectory = Join-Path $env:LOCALAPPDATA 'OwlLawyer\Logs'
    New-Item -ItemType Directory -Path $logDirectory -Force | Out-Null
    $log = Join-Path $logDirectory 'setup.log'
    Start-Transcript -Path $log -Force | Out-Null
    try {
        if (-not $CheckOnly) { New-Item -ItemType Directory -Path $InstallDir -Force | Out-Null }
        Initialize-SovaRuntime
        Stop-Transcript | Out-Null
        exit 0
    } catch {
        $message = $_.Exception.Message
        Write-Host "SETUP FAILED: $message"
        if ($ReportFile) { [IO.File]::WriteAllText($ReportFile, $message, [Text.UTF8Encoding]::new($false)) }
        Stop-Transcript | Out-Null
        exit 1
    }
}
