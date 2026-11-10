# Builds and publishes a Lab-I release on GitHub so support staff can download it
# from the Releases page.
#
#   .\tools\release.ps1 -DryRun                      build + show notes, change nothing
#   .\tools\release.ps1                              ComSet, HostForms_SA, Interface.exe
#   .\tools\release.ps1 -Projects HostForms_SA       one component (tag HostForms_SA-v<ver>)
#
# Steps:
#   1. Requires a clean main.
#   2. Builds with tools\build.ps1.
#   3. Commits the .vbp version bumps VB6 makes.
#   4. Tags v<Interface version>, or <Component>-v<version> when Interface.exe is
#      not included.
#   5. Pushes.
#   6. Creates the GitHub release with the built files and support-facing notes.
#
# Also used by .github/workflows/release.yml on the self-hosted VB6 runner.
param(
    [string[]]$Projects = @('ComSet', 'HostForms_SA', 'Interface'),
    [switch]$DryRun,
    [string]$Repo = 'strugz/Interface-Main-Application',
    [string]$TestedNote = 'Built from source on the VB6 build PC. Not yet tested at a site with a live analyzer.'
)

$ErrorActionPreference = 'Stop'
$root = Split-Path $PSScriptRoot -Parent
Set-Location $root
if ($Projects.Count -eq 1 -and $Projects[0] -match ',') { $Projects = $Projects[0] -split '\s*,\s*' }

function Invoke-Git { & git.exe @args; if ($LASTEXITCODE -ne 0) { throw "git $($args -join ' ') failed" } }
function Fail($msg) { Write-Host "ERROR: $msg" -ForegroundColor Red; exit 1 }

$outputs = [ordered]@{
    ComSet       = @{ file = 'New Lab-I System v2\ComSet.dll';       progid = 'Comset.ComSettings' }
    HostForms_SA = @{ file = 'New Lab-I System v2\HostForms_SA.dll'; progid = 'HostForms_SA.ClassForms' }
    HostForms_RE = @{ file = 'HostFomrs_RE\HostForms_RE.dll';        progid = 'HostForms_RE.ClassForms' }
    Interface    = @{ file = 'New Lab-I System v2\Interface.exe';    progid = $null }
}
foreach ($p in $Projects) { if (-not $outputs.Contains($p)) { Fail "unknown project '$p' (use: $($outputs.Keys -join ', '))" } }

# ---- 1. clean main -----------------------------------------------------------
$branch = (& git rev-parse --abbrev-ref HEAD).Trim()
if ($branch -ne 'main') { Fail "releases are made from main (current branch: $branch)" }
$dirty = & git status --porcelain
if ($dirty) { Fail "working tree is not clean; commit or stash first:`n$($dirty -join "`n")" }
if (-not $DryRun) {
    Invoke-Git fetch --quiet --tags origin
    $behind = (& git rev-list --count HEAD..origin/main).Trim()
    if ($behind -ne '0') { Fail "main is $behind commit(s) behind origin/main; pull first" }
}
$eap = $ErrorActionPreference; $ErrorActionPreference = 'Continue'   # PS 5.1: stderr redirect + Stop throws
$prevTag = (& git describe --tags --abbrev=0 2>$null)
$describeExit = $LASTEXITCODE
$ErrorActionPreference = $eap
if ($describeExit -ne 0) { $prevTag = $null }

# ---- 2. binary-compatibility reference for ComSet ----------------------------
# ComSet.vbp has CompatibleMode=1 against ComSet\ComSet.dll, which is not in git.
# Use the last released ComSet.dll when it is missing (for example on a fresh
# runner checkout).
if ($Projects -contains 'ComSet' -and -not (Test-Path 'ComSet\ComSet.dll')) {
    Write-Host 'ComSet\ComSet.dll missing: downloading the last released one for binary compatibility'
    & gh release download --repo $Repo --pattern 'ComSet.dll' --dir 'ComSet' --clobber
    if ($LASTEXITCODE -ne 0 -or -not (Test-Path 'ComSet\ComSet.dll')) { Fail 'could not get ComSet.dll for binary compatibility' }
}

# ---- 3. remember the current COM registrations (32-bit view) -----------------
# VB6 /make registers what it builds. On the runner that would point Lab-I on
# this PC at the runner's workspace, so the previous registrations are restored
# after the build.
function Get-ComPath($progid) {
    $hkcr = [Microsoft.Win32.RegistryKey]::OpenBaseKey('ClassesRoot', 'Registry32')
    $k = $hkcr.OpenSubKey("$progid\CLSID"); if (-not $k) { return $null }
    $clsid = $k.GetValue(''); $k.Close()
    $k = $hkcr.OpenSubKey("CLSID\$clsid\InprocServer32"); if (-not $k) { return $null }
    $path = $k.GetValue(''); $k.Close(); return $path
}
$previous = @{}
foreach ($p in $Projects) {
    $id = $outputs[$p].progid
    if ($id) { $previous[$p] = Get-ComPath $id }
}

# ---- 4. build -----------------------------------------------------------------
& (Join-Path $PSScriptRoot 'build.ps1') -Projects $Projects
$buildFailed = $LASTEXITCODE

foreach ($p in $previous.Keys) {
    $old = $previous[$p]
    $new = (Resolve-Path $outputs[$p].file -ErrorAction SilentlyContinue).Path
    if ($old -and (Test-Path $old) -and $new -and ($old -ne $new)) {
        # regsvr32 is a GUI program (PowerShell would not wait for it), and HKLM
        # needs elevation: VB6.EXE runs as admin on the build PC, so do the same.
        $reg = Start-Process "$env:SystemRoot\SysWOW64\regsvr32.exe" -ArgumentList '/s', "`"$old`"" -Verb RunAs -Wait -PassThru
        if ($reg.ExitCode -eq 0 -and (Get-ComPath $outputs[$p].progid) -eq $old) {
            Write-Host "Restored COM registration of $p to $old"
        } else {
            $msg = "Could not restore COM registration of $p to $old (regsvr32 exit $($reg.ExitCode)); run as admin: regsvr32 `"$old`""
            Write-Host "WARNING: $msg" -ForegroundColor Yellow
            if ($env:GITHUB_ACTIONS) { Write-Host "::warning::$msg" }
        }
    }
}
if ($buildFailed -ne 0) { Fail "$buildFailed project(s) failed to build" }

$files = foreach ($p in $Projects) {
    $f = Get-Item $outputs[$p].file
    [pscustomobject]@{
        Name    = $f.Name
        Path    = $f.FullName
        Version = ($f.VersionInfo.FileMajorPart, $f.VersionInfo.FileMinorPart, $f.VersionInfo.FilePrivatePart) -join '.'
        Size    = $f.Length
        Hash    = (Get-FileHash $f.FullName -Algorithm SHA256).Hash
        Project = $p
    }
}

# ---- 5. tag name and notes ------------------------------------------------------
$iface = $files | Where-Object Project -eq 'Interface'
if ($iface) { $tag = "v$($iface.Version)"; $title = "Lab-I Interface $($iface.Version)" }
else { $first = $files[0]; $tag = "$($first.Project)-v$($first.Version)"; $title = "$($first.Project) $($first.Version)" }
$summary = ($files | ForEach-Object { "$($_.Name) $($_.Version)" }) -join ', '

if (& git tag --list $tag) { Fail "tag $tag already exists locally" }
if (-not $DryRun) {
    & git ls-remote --exit-code --tags origin "refs/tags/$tag" *> $null
    if ($LASTEXITCODE -eq 0) { Fail "tag $tag already exists on GitHub" }
}

$range = if ($prevTag) { "$prevTag..HEAD" } else { 'HEAD' }
$changes = & git log $range --no-merges --format='- %s' | Where-Object { $_ -notmatch '^- Build ' }
if (-not $changes) { $changes = @('- Rebuild only; no source changes since the previous release.') }

$nl = "`r`n"
$notes = @(
    "Built from tag ``$tag`` with ``tools/release.ps1``.", '',
    '| File | Version | SHA-256 |', '|---|---|---|'
) + ($files | ForEach-Object { "| ``$($_.Name)`` | $($_.Version) | ``$($_.Hash)`` |" }) + @(
    '', "## Changes since $(if ($prevTag) { $prevTag } else { 'the start of the repository' })", ''
) + $changes + @(
    '', '## Testing', '', $TestedNote, '',
    '## Install / upgrade', '',
    '1. Close Lab-I and back up the files you are replacing in the Lab-I folder.',
    '2. Copy the downloaded files into the Lab-I folder.',
    '3. From an **administrator** prompt, register each DLL with the 32-bit regsvr32:', '   ```'
) + ($files | Where-Object { $_.Name -like '*.dll' } | ForEach-Object { "   C:\Windows\SysWOW64\regsvr32 ""<Lab-I folder>\$($_.Name)""" }) + @(
    '   ```',
    '4. Start Lab-I and check **Help > About** for the new version.', '',
    'Not included: instrument plugins (they come from the LIS Driver project), runtime prerequisites (Crystal Reports 8, OCX controls, ADO, DAO) and databases.'
)
$notesFile = Join-Path $env:TEMP "lab-i-release-$tag.md"
[IO.File]::WriteAllText($notesFile, ($notes -join $nl), (New-Object Text.UTF8Encoding $false))

Write-Host ''
Write-Host "Release $tag : $summary" -ForegroundColor Cyan
Get-Content $notesFile | ForEach-Object { Write-Host "  $_" }

if ($DryRun) {
    Invoke-Git checkout -- '*.vbp'
    Write-Host ''
    Write-Host 'Dry run: version bumps discarded; nothing committed, tagged or published.' -ForegroundColor Yellow
    exit 0
}

# ---- 6. commit, tag, push, publish ------------------------------------------------
if (-not (& git config user.name)) { Invoke-Git config user.name 'Lab-I build bot' }
if (-not (& git config user.email)) { Invoke-Git config user.email 'lab-i-build-bot@users.noreply.github.com' }
Invoke-Git add -- '*.vbp'
Invoke-Git commit --quiet -m "Build $summary" -m "Version bumps made by VB6 /make for release $tag."
Invoke-Git tag -a $tag -m "$title" -m "$summary"
Invoke-Git push --quiet origin main
Invoke-Git push --quiet origin $tag

$assets = $files | ForEach-Object { $_.Path }
& gh release create $tag @assets --repo $Repo --title $title --notes-file $notesFile --verify-tag
if ($LASTEXITCODE -ne 0) { Fail "gh release create failed (tag $tag is pushed; rerun 'gh release create' by hand)" }

# ---- 7. verify the uploads ---------------------------------------------------------
$remote = & gh release view $tag --repo $Repo --json assets,url | ConvertFrom-Json
foreach ($f in $files) {
    $a = $remote.assets | Where-Object name -eq $f.Name
    if (-not $a -or $a.size -ne $f.Size) { Fail "uploaded $($f.Name) does not match the local build" }
}
Write-Host ''
Write-Host "Published: $($remote.url)" -ForegroundColor Green
if ($env:GITHUB_STEP_SUMMARY) {
    Add-Content $env:GITHUB_STEP_SUMMARY "### Released [$title]($($remote.url))`n`n$summary"
}
