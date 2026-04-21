param(
  [Parameter(Mandatory = $false)][string]$AsciiPrefix = "dis_",
  [Parameter(Mandatory = $false)][string]$Extension = ".docx"
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

$root = (Get-Location).Path

$candidates =
  Get-ChildItem -LiteralPath $root -Filter ("*" + $Extension) -File |
  Where-Object { $_.Name.StartsWith($AsciiPrefix) } |
  ForEach-Object {
    $m = [Regex]::Match($_.BaseName, "^(?<base>.+)_(?<num>\d+)$")
    if ($m.Success) {
      [pscustomobject]@{
        File = $_
        Base = $m.Groups["base"].Value
        Num  = [int]$m.Groups["num"].Value
      }
    }
  } |
  Where-Object { $_ -ne $null }

if (-not $candidates -or @($candidates).Count -eq 0) {
  throw "No versioned docx found (expected something like 'dis_*_<N>${Extension}')."
}

$max = ($candidates | Sort-Object Num -Descending | Select-Object -First 1)
$nextNum = $max.Num + 1

$src = $max.File.FullName
$dstName = "{0}_{1}{2}" -f $max.Base, $nextNum, $Extension
$dst = Join-Path $root $dstName

Copy-Item -LiteralPath $src -Destination $dst -Force

@{
  source = $src
  output = $dst
  version_from = $max.Num
  version_to = $nextNum
  base = $max.Base
} | ConvertTo-Json -Depth 3

