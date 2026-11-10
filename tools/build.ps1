# Builds the Lab-I components with VB6, in dependency order.
#   .\tools\build.ps1                      ComSet, HostForms_SA, Interface.exe
#   .\tools\build.ps1 -Projects ComSet     only the named ones (ComSet, HostForms_SA, HostForms_RE, Interface)
# Each project's output goes to its Path32 folder (New Lab-I System v2 for the
# DLLs; Interface.exe beside its .vbp). VB6 /make also registers the DLLs.
# VB6 bumps RevisionVer in the .vbp on every make: commit that change.
param([string[]]$Projects = @('ComSet', 'HostForms_SA', 'Interface'))

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
$vb6 = @(
    "${env:ProgramFiles(x86)}\Microsoft Visual Studio\VB98\VB6.EXE",
    "$env:ProgramFiles\Microsoft Visual Studio\VB98\VB6.EXE"
) | Where-Object { Test-Path $_ } | Select-Object -First 1
if (-not $vb6) { throw 'VB6.EXE not found (Visual Basic 6.0 must be installed).' }

$all = [ordered]@{
    ComSet       = @{ vbp = 'ComSet\ComSet.vbp';                     out = 'New Lab-I System v2\ComSet.dll' }
    HostForms_SA = @{ vbp = 'HostForms_SA v2\HostForms_SA.vbp';      out = 'New Lab-I System v2\HostForms_SA.dll' }
    HostForms_RE = @{ vbp = 'HostFomrs_RE\HostForms_RE.vbp';         out = 'HostFomrs_RE\HostForms_RE.dll' }
    Interface    = @{ vbp = 'New Lab-I System v2\HostInterface.vbp'; out = 'New Lab-I System v2\Interface.exe' }
}

$failed = 0
foreach ($name in $all.Keys) {
    if ($Projects -notcontains $name) { continue }
    $p = $all[$name]
    $vbp = Join-Path $root $p.vbp
    $dll = Join-Path $root $p.out
    $log = Join-Path $env:TEMP "lab-i-build-$name.log"
    Remove-Item $log -ErrorAction SilentlyContinue
    $before = if (Test-Path $dll) { (Get-Item $dll).LastWriteTime } else { [datetime]::MinValue }

    Write-Host "Building $name ..." -NoNewline
    $proc = Start-Process $vb6 -ArgumentList "/make `"$vbp`" /out `"$log`"" -Wait -PassThru
    $ok = $proc.ExitCode -eq 0 -and (Test-Path $dll) -and (Get-Item $dll).LastWriteTime -gt $before
    if ($ok) {
        Write-Host (' OK  v{0}' -f (Get-Item $dll).VersionInfo.FileVersion) -ForegroundColor Green
    } else {
        Write-Host ' FAILED' -ForegroundColor Red
        $failed++
    }
    if (Test-Path $log) { Get-Content $log | ForEach-Object { Write-Host "   $_" } }
}
exit $failed
