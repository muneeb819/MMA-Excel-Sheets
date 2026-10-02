<#
.SYNOPSIS
    One-shot setup for SheetCraft.

.DESCRIPTION
    Installs dependencies, type-checks, builds the app and (optionally) produces
    the Windows installer.

.EXAMPLE
    .\setup.ps1
    .\setup.ps1 -Package
    .\setup.ps1 -SkipInstall    # dependencies are already present
#>
[CmdletBinding()]
param(
    [switch]$Package,
    [switch]$SkipInstall,
    [switch]$SkipTests
)

$ErrorActionPreference = 'Stop'
$root = $PSScriptRoot
Set-Location $root

function Write-Step($message) {
    Write-Host ''
    Write-Host "==> $message" -ForegroundColor Cyan
}

# npm is blocked by script policy on some machines; npm.cmd always works.
$npm = if (Get-Command npm.cmd -ErrorAction SilentlyContinue) { 'npm.cmd' } else { 'npm' }

Write-Host 'SheetCraft setup' -ForegroundColor Green
Write-Host "  folder: $root"

if (-not $SkipInstall) {
    Write-Step 'Installing dependencies'
    & $npm install --no-audit --no-fund
    if ($LASTEXITCODE -ne 0) { throw 'npm install failed' }

    # npm 11 blocks postinstall scripts by default; esbuild and electron need
    # theirs to fetch their binaries.
    $pending = & node -e "try{const p=require('./package.json');const l=require('./package-lock.json');const n=Object.keys(l.packages||{}).filter(k=>/esbuild|^node_modules\/electron$/.test(k));console.log(n.length?'':'none')}catch(e){console.log('none')}"
    if ($pending -eq '') {
        Write-Step 'Approving required install scripts'
        & $npm approve-scripts esbuild electron 2>$null | Out-Null
        & $npm rebuild esbuild electron 2>$null | Out-Null
    }
}

if (-not $SkipTests) {
    Write-Step 'Type checking'
    & $npm run typecheck
    if ($LASTEXITCODE -ne 0) { throw 'Type check failed' }

    Write-Step 'Running engine self-test'
    & $npm run selftest
    if ($LASTEXITCODE -ne 0) { throw 'Self-test failed' }
}

Write-Step 'Building the app'
& $npm run build
if ($LASTEXITCODE -ne 0) { throw 'Build failed' }

if ($Package) {
    Write-Step 'Packaging the Windows installer'
    & $npm run dist
    if ($LASTEXITCODE -ne 0) { throw 'Packaging failed' }

    $release = Join-Path $root 'release'
    Write-Host ''
    Write-Host 'Installer ready:' -ForegroundColor Green
    Get-ChildItem $release -Filter '*.exe' -ErrorAction SilentlyContinue |
        ForEach-Object { Write-Host "  $($_.FullName)  ($([math]::Round($_.Length / 1MB, 1)) MB)" }
} else {
    Write-Host ''
    Write-Host 'Done. Start the app with:' -ForegroundColor Green
    Write-Host '  npm run desktop      (desktop window)'
    Write-Host '  npm run dev          (browser, http://localhost:5173)'
    Write-Host ''
    Write-Host 'To produce an installer run:  npm run dist' -ForegroundColor DarkGray
}
