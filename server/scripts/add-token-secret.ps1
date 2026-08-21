# Add ADMIN_TOKEN_SECRET to .env  --  ASCII ONLY on purpose.
#
# PowerShell 5.1 reads a BOM-less .ps1 as ANSI. A script containing Japanese
# gets mangled and once wrote an EMPTY .env (2026-08-16). Keep this file ASCII.
#
# What it does:
#   1. back up .env
#   2. read the secret from a one-line file
#   3. replace the key if present, otherwise append it
#   4. verify the file did not become empty
#   5. delete the source file
#
# Usage (from server/):
#   powershell -ExecutionPolicy Bypass -File scripts\add-token-secret.ps1 -From "$env:USERPROFILE\secret.txt"

param(
    [Parameter(Mandatory=$true)][string]$From
)

$ErrorActionPreference = "Stop"
$envPath = Join-Path $PSScriptRoot "..\.env"
$envPath = (Resolve-Path $envPath).Path

if (-not (Test-Path $From))    { Write-Output "NG: source file not found: $From"; exit 1 }
if (-not (Test-Path $envPath)) { Write-Output "NG: .env not found: $envPath";     exit 1 }

$before = (Get-Item $envPath).Length
if ($before -lt 10) { Write-Output "NG: .env looks empty already ($before bytes). stop."; exit 1 }

# ---- back up first ----
$stamp  = Get-Date -Format "yyyyMMdd-HHmmss"
$backup = "$envPath.backup-$stamp"
Copy-Item $envPath $backup
Write-Output ("backed up -> " + (Split-Path -Leaf $backup) + " (" + $before + " bytes)")

# ---- read the secret ----
$line = (Get-Content $From -Raw).Trim()
if ($line -notmatch '^ADMIN_TOKEN_SECRET=(.+)$') {
    Write-Output "NG: the file must contain exactly one line: ADMIN_TOKEN_SECRET=..."
    exit 1
}
$value = $matches[1].Trim().Trim('"').Trim("'")
if ($value.Length -lt 8) { Write-Output "NG: value looks too short."; exit 1 }
Write-Output ("secret read. length = " + $value.Length + " characters (value not shown)")

# ---- replace or append ----
$lines   = Get-Content $envPath -Encoding UTF8
$found   = $false
$updated = @()
foreach ($l in $lines) {
    if ($l -match '^\s*ADMIN_TOKEN_SECRET\s*=') {
        $updated += ("ADMIN_TOKEN_SECRET=" + $value)
        $found = $true
    } else {
        $updated += $l
    }
}
if (-not $found) {
    $updated += ""
    $updated += "# Used to verify the tokens issued by GAS. DO NOT CHANGE THIS VALUE."
    $updated += "# Changing it invalidates every token already on customers' devices."
    $updated += ("ADMIN_TOKEN_SECRET=" + $value)
}

# write to a temp file first, then move -- never write .env directly
$tmp = "$envPath.tmp"
$utf8 = New-Object System.Text.UTF8Encoding($false)
[System.IO.File]::WriteAllLines($tmp, $updated, $utf8)

$after = (Get-Item $tmp).Length
if ($after -lt $before) {
    Write-Output ("NG: new file is smaller (" + $after + " < " + $before + "). aborting, .env untouched.")
    Remove-Item $tmp -Force
    exit 1
}
Move-Item $tmp $envPath -Force

$final = (Get-Item $envPath).Length
Write-Output ("done. .env is now " + $final + " bytes  (was " + $before + ")")
if ($found) { Write-Output "the key already existed and was replaced." }
else        { Write-Output "the key was appended." }

# ---- clean up ----
Remove-Item $From -Force
Write-Output "source file deleted."
Write-Output ""
Write-Output "next: docker compose up -d      (the .env is read at startup)"
