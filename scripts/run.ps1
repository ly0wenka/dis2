param(
  [Parameter(Mandatory = $false)][switch]$ForceDirty,
  [Parameter(Mandatory = $false)][switch]$DeleteStruckText,
  [Parameter(Mandatory = $false)][switch]$NoGit,
  [Parameter(Mandatory = $false)][switch]$NoPush,
  [Parameter(Mandatory = $false)][string]$InputDocx,
  [Parameter(Mandatory = $false)][int]$TargetVersion,
  [Parameter(Mandatory = $false)][switch]$OnlyRewrite44
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

function Resolve-InputDocxPath {
  param([Parameter(Mandatory = $true)][string]$Path)
  return (Resolve-Path -LiteralPath $Path).Path
}

# Determine output docx + version
$bump = $null
$outDocx = $null

if ($TargetVersion) {
  if (-not $InputDocx) { throw "When -TargetVersion is set, -InputDocx is required." }
  $outDocx = Join-Path $repoRoot ("dis_Кондратов_{0}.docx" -f $TargetVersion)
  $bump = [pscustomobject]@{ version_to = $TargetVersion; output = $outDocx }
} else {
  $bumpJson = & "$PSScriptRoot\\bump_docx_version.ps1"
  $bump = $bumpJson | ConvertFrom-Json
  $outDocx = $bump.output
}

# Backup existing output file if present
$ts = Get-Date -Format "yyyyMMdd_HHmmss"
$leafBase = [System.IO.Path]::GetFileNameWithoutExtension($outDocx)
if (Test-Path -LiteralPath $outDocx) {
  Copy-Item -LiteralPath $outDocx -Destination (Join-Path $repoRoot ("tmp\\backups\\{0}_pre_overwrite_{1}.docx" -f $leafBase, $ts)) -Force
}

$report = Join-Path $repoRoot ("tmp\\reports\\report_{0}.json" -f $bump.version_to)

if ($OnlyRewrite44) {
  if (-not $InputDocx) { throw "When -OnlyRewrite44 is set, -InputDocx is required." }
  & "$PSScriptRoot\\rewrite_44.ps1" `
    -InputDocx $InputDocx `
    -OutputDocx $outDocx `
    -ReportPath $report
} else {
  # Default path (legacy)
  # If InputDocx is provided, copy it into the repo root first so bumping uses it.
  if ($InputDocx) {
    $in = Resolve-InputDocxPath -Path $InputDocx
    $name = [System.IO.Path]::GetFileName($in)
    $dest = Join-Path $repoRoot $name
    if (Test-Path -LiteralPath $dest) {
      $destBase = [System.IO.Path]::GetFileNameWithoutExtension($dest)
      Copy-Item -LiteralPath $dest -Destination (Join-Path $repoRoot ("tmp\\backups\\{0}_pre_input_{1}.docx" -f $destBase, $ts)) -Force
    }
    Copy-Item -LiteralPath $in -Destination $dest -Force
  }

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
}

if (-not $NoGit) {
  git add -- (Split-Path -Leaf $outDocx) scripts
  if ($OnlyRewrite44) {
    git commit -m ("v{0}: rewrite §4.4 (style-safe) + refs in 4th sentence" -f $bump.version_to)
  } else {
    git commit -m ("v{0}: normalize text + 4.4 cleanup; scripts" -f $bump.version_to)
  }
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

