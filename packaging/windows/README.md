# Windows installer

`installer.iss` builds a per-user Inno Setup wizard, `SOVA-Setup-0.2.1-x64.exe`.
Targets x64 Windows 10 1803+ and Windows 11. The installer is online: missing
third-party packages and Python wheels are downloaded during setup. The app
itself uses no network services.

## Build on Windows

Install Python with Tkinter and Inno Setup 6.3+ (ISCC on PATH or in its default
installation folder), then run from the repository root:

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File packaging/windows/build.ps1
```

Output: `dist/SOVA-Setup-0.2.1-x64.exe` and its SHA-256 file.
`-Version 0.2.1` overrides the version. The build downloads the official Russian
OCR model from `tesseract-ocr/tessdata_fast` tag `4.1.0`, checks its pinned SHA-256,
and wraps the existing owl PNG in a Windows ICO container.
The OCR model's Apache 2.0 license is included under `licenses/`.

The `Windows installer` GitHub Actions workflow runs setup tests on Windows,
builds the installer, and uploads it as `SOVA-Windows-installer`.

## Build on macOS/Linux

Prepare assets and use the pinned compiler container:

```bash
python3 packaging/windows/build_assets.py
docker run --rm --platform linux/amd64 -v "$PWD:/work" \
  amake/innosetup@sha256:e003376ba818547275fe10c95e2a29be0f2d12d45e9eb8f205b6672dc5685bb1 \
  packaging/windows/installer.iss
```

This produces the same executable. A cross-build does not validate real Windows
installation, UAC behavior, or third-party package installation. Inno Setup's
license terms apply to the compiler; configure an appropriate licensed compiler
for commercial release builds.

## Setup behavior

- Install only into the current user's `%LOCALAPPDATA%\Programs\OwlLawyer`.
- Check for compatible 64-bit Python 3.11–3.14, Tkinter, pip, and venv. Discover
  registry, known installation directories, and PATH; skip Store stub aliases.
- With the user's selected task, install missing tools through WinGet using exact
  package IDs: `Python.Python.3.14`, `UB-Mannheim.TesseractOCR`, and
  `oschwartz10612.Poppler`. WinGet verifies package downloads against its catalog.
- On Windows 10 1803 (build 17134), or when WinGet is missing, use the pinned
  vendor downloads in `dependency-manifest.json`. WinGet itself requires build
  17763+, so do not ask users of 1803 to install it. Verify SHA-256 before running
  downloaded installers, reuse valid cached downloads, install the Microsoft
  runtime when needed, and unpack Poppler privately under `tools/poppler/`.
- The direct installer versions and checksums come from Microsoft's WinGet
  catalog; the files are downloaded from Python.org, UB Mannheim's GitHub
  releases, Poppler's Windows releases, The Document Foundation, and Microsoft.
- Offer `TheDocumentFoundation.LibreOffice` separately; reuse it when detected.
- Install pinned Python wheels only in `runtime/`. Never change global pip packages.
- Include the checksum-verified Russian model in `tessdata/`; verify Tesseract
  reports `rus`. Store tool locations in `runtime-tools.json`, which the app loads
  into its own process environment. Do not modify system PATH.
- Abort before installing app files on prerequisite failure. WinGet/other installed
  prerequisites and a partially prepared private runtime may remain for a retry.
- Log dependency setup to `%LOCALAPPDATA%\OwlLawyer\Logs\setup.log`.
- Preserve all user workbooks, scans, generated documents, incoming documents,
  and editable templates during update/uninstall. Remove only the tracked program
  files, shortcuts, private runtime, and runtime configuration.

`install-windows.bat` is a source-distribution fallback. It uses the same dependency
script and asks before installing missing components/LibreOffice. It does not
register an uninstaller; use the compiled wizard for normal distribution.

## Verification

```powershell
python -m unittest discover -s tests -p test_windows_runtime.py -v
powershell -NoProfile -ExecutionPolicy Bypass -File tests/test_windows_setup.ps1
powershell -NoProfile -ExecutionPolicy Bypass -File tests/test_windows_downloads.ps1
```

The PowerShell tests mock package installation; they do not install software.
They exercise Python detection, native stderr/exit codes, consent, component
reuse, optional LibreOffice, missing tools, read-only checks, and failures.
Download tests also cover build 17134 bypassing WinGet, missing-WinGet fallback,
cached file verification, corrupt downloads, cancelled installers, and reboots.

Before distributing a release, test a clean Windows VM, existing Python without
Tkinter, tools installed outside PATH, rejected UAC, offline failure/retry, and
upgrading/uninstalling with real documents present. The first build is unsigned;
add Authenticode signing with the publisher's certificate for public releases.

Silent installation is supported by Inno Setup. Default tasks install missing
components and LibreOffice; select tasks explicitly for managed deployments:

```powershell
.\SOVA-Setup-0.2.1-x64.exe /VERYSILENT /SUPPRESSMSGBOXES /NORESTART /TASKS="installmissing,libreoffice,desktopicon"
```

Third-party installers can still request UAC. With all dependencies preinstalled,
use `/TASKS=""` to omit third-party installations. It still prepares the private
Python environment and may need the internet for Python wheels.
