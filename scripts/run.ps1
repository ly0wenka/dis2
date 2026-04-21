param(
  [Parameter(Mandatory = $false)][switch]$ForceDirty,
  [Parameter(Mandatory = $false)][switch]$DeleteStruckText,
  [Parameter(Mandatory = $false)][switch]$NoGit,
  [Parameter(Mandatory = $false)][switch]$NoPush,
  # Optional: path to an externally edited base DOCX (e.g. your manual 111).
  [Parameter(Mandatory = $false)][string]$InputDocx
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$repoRoot = (Get-Location).Path

$dirty = @(& git status --porcelain).Count -gt 0
if ($dirty -and -not $ForceDirty) {
  throw "Git working tree is dirty. Commit/stash or re-run with -ForceDirty."
}

New-Item -ItemType Directory -Path (Join-Path $repoRoot "tmp\\backups") -Force | Out-Null
New-Item -ItemType Directory -Path (Join-Path $repoRoot "tmp\\reports") -Force | Out-Null

if ($InputDocx) {
  $in = (Resolve-Path -LiteralPath $InputDocx).Path
  $name = [System.IO.Path]::GetFileName($in)
  $dest = Join-Path $repoRoot $name
  $ts = Get-Date -Format "yyyyMMdd_HHmmss"
  if (Test-Path -LiteralPath $dest) {
    $destBase = [System.IO.Path]::GetFileNameWithoutExtension($dest)
    Copy-Item -LiteralPath $dest -Destination (Join-Path $repoRoot ("tmp\\backups\\{0}_pre_input_{1}.docx" -f $destBase, $ts)) -Force
  }
  Copy-Item -LiteralPath $in -Destination $dest -Force
}

$bumpJson = & "$PSScriptRoot\\bump_docx_version.ps1"
$bump = $bumpJson | ConvertFrom-Json

$outDocx = $bump.output

# Backup before modifying
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$leafBase = [System.IO.Path]::GetFileNameWithoutExtension($outDocx)
$backupOut = Join-Path $repoRoot ("tmp\\backups\\{0}_pre_norm_{1}.docx" -f $leafBase, $ts)
Copy-Item -LiteralPath $outDocx -Destination $backupOut -Force

$report = Join-Path $repoRoot ("tmp\\reports\\report_{0}.json" -f $bump.version_to)
$strikeMode = "KeepTextRemoveStrike"
if ($DeleteStruckText) { $strikeMode = "DeleteText" }

& "$PSScriptRoot\\normalize_docx.ps1" `
  -InputDocx $outDocx `
  -OutputDocx $outDocx `
  -Scope Whole `
  -Fix44 `
  -Remove44Formulas `
  -NormalizeBracketCitations `
  -RemoveRefTokens `
  -StrikeHandling $strikeMode `
  -ReportPath $report

if (-not $NoGit) {
  git add -- (Split-Path -Leaf $outDocx) scripts
  git commit -m ("v{0}: normalize text + 4.4 cleanup; scripts" -f $bump.version_to)
  git tag $bump.version_to
  if (-not $NoPush) {
    git push --follow-tags
  }
}

if ($NoGit) {
  Write-Host ("DONE (no git): created {0}" -f (Split-Path -Leaf $outDocx))
} elseif ($NoPush) {
  Write-Host ("DONE (no push): created tag {0} and file {1}" -f $bump.version_to, (Split-Path -Leaf $outDocx))
} else {
  Write-Host ("DONE: pushed tag {0} and file {1}" -f $bump.version_to, (Split-Path -Leaf $outDocx))
}

