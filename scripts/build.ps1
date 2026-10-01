param([string]$Python = 'python')
$ErrorActionPreference = 'Stop'
$taskProjectDir = Split-Path -Parent $PSScriptRoot
$taskBuildDir = Join-Path $taskProjectDir 'work\build'
$taskDistDir = Join-Path $taskProjectDir 'dist'
New-Item -ItemType Directory -Force -Path $taskBuildDir, $taskDistDir | Out-Null
& $Python -m PyInstaller --noconfirm --onefile --windowed --name DocumentScanCleanup --hidden-import win32timezone --collect-data pikepdf --paths $taskProjectDir --distpath $taskDistDir --workpath $taskBuildDir --specpath $taskBuildDir --exclude-module pandas --exclude-module numpy --exclude-module pyarrow --exclude-module scipy --exclude-module torch --exclude-module matplotlib --exclude-module cv2 --exclude-module IPython --exclude-module sklearn (Join-Path $taskProjectDir 'document_scan_cleanup.py')
if ($LASTEXITCODE -ne 0) { throw 'Build failed.' }
& $Python (Join-Path $PSScriptRoot 'audit_public.py') --exe (Join-Path $taskDistDir 'DocumentScanCleanup.exe')
if ($LASTEXITCODE -ne 0) { throw 'Public executable audit failed.' }
Write-Output (Join-Path $taskDistDir 'DocumentScanCleanup.exe')
