param([string]$Python = 'python')
$ErrorActionPreference = 'Stop'
& $Python (Join-Path $PSScriptRoot 'package.py')
if ($LASTEXITCODE -ne 0) { throw 'Packaging failed.' }
