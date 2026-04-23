param(
  [Parameter(Mandatory = $true)][string]$InputDocx,
  [Parameter(Mandatory = $true)][string]$OutputDocx,
  [Parameter(Mandatory = $false)][ValidateSet("lingva", "mymemory")][string]$Provider = "lingva",
  [Parameter(Mandatory = $false)][string]$SourceLang = "uk",
  [Parameter(Mandatory = $false)][string]$TargetLang = "en",
  [Parameter(Mandatory = $false)][string]$Email = "",
  [Parameter(Mandatory = $false)][ValidateRange(50, 2000)][int]$MaxChunkChars = 450,
  [Parameter(Mandatory = $false)][ValidateRange(0, 2000)][int]$SleepMs = 200,
  [Parameter(Mandatory = $false)][string]$CachePath = "",
  [Parameter(Mandatory = $false)][string]$ReportPath = ""
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

. "$PSScriptRoot\\docx_lib.ps1"

function Ensure-ParentDir {
  param([Parameter(Mandatory = $true)][string]$Path)
  $parent = Split-Path -Parent $Path
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent -Force | Out-Null
  }
}

function Has-Cyrillic {
  param([Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text)
  return ($Text -match "[\\p{IsCyrillic}]")
}

function Split-TextForMyMemory {
  param(
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][int]$MaxLen
  )
  if ($Text.Length -le $MaxLen) { return @($Text) }

  $chunks = New-Object System.Collections.Generic.List[string]
  $i = 0
  while ($i -lt $Text.Length) {
    $take = [Math]::Min($MaxLen, $Text.Length - $i)
    $window = $Text.Substring($i, $take)

    # Prefer splitting on sentence boundaries, then spaces.
    $cut = -1
    foreach ($rx in @(
      [regex]::new("(?s).*[\.!\?;:]\s", [System.Text.RegularExpressions.RegexOptions]::None),
      [regex]::new("(?s).*\s", [System.Text.RegularExpressions.RegexOptions]::None)
    )) {
      $m = $rx.Match($window)
      if ($m.Success) {
        $cut = $m.Length
        break
      }
    }
    if ($cut -le 0) { $cut = $window.Length }

    $chunks.Add($Text.Substring($i, $cut))
    $i += $cut
  }
  return $chunks.ToArray()
}

function Load-Cache {
  param([Parameter(Mandatory = $true)][string]$Path)
  if (-not (Test-Path -LiteralPath $Path)) { return @{} }
  $raw = Get-Content -LiteralPath $Path -Raw
  if (-not $raw.Trim()) { return @{} }
  $obj = $raw | ConvertFrom-Json
  $ht = @{}
  foreach ($p in $obj.PSObject.Properties) { $ht[$p.Name] = [string]$p.Value }
  return $ht
}

function Save-Cache {
  param(
    [Parameter(Mandatory = $true)][hashtable]$Cache,
    [Parameter(Mandatory = $true)][string]$Path
  )
  Ensure-ParentDir -Path $Path
  # Convert hashtable to a stable JSON object.
  $ordered = [ordered]@{}
  foreach ($k in ($Cache.Keys | Sort-Object)) { $ordered[$k] = $Cache[$k] }
  ($ordered | ConvertTo-Json -Depth 3) | Set-Content -LiteralPath $Path -Encoding UTF8
}

function Invoke-MyMemoryTranslate {
  param(
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Dst,
    [Parameter(Mandatory = $false)][string]$Email,
    [Parameter(Mandatory = $false)][int]$TimeoutSec = 30
  )
  $q = [uri]::EscapeDataString($Text)
  $url = "https://api.mymemory.translated.net/get?q=$q&langpair=$Src|$Dst"
  if ($Email) { $url += "&de=$([uri]::EscapeDataString($Email))" }
  $resp = Invoke-RestMethod -Uri $url -TimeoutSec $TimeoutSec
  if (-not $resp -or $resp.responseStatus -ne 200) {
    $detail = ""
    if ($resp -and $resp.responseDetails) { $detail = [string]$resp.responseDetails }
    throw "MyMemory translate failed (status=$($resp.responseStatus)) $detail"
  }
  return [System.Net.WebUtility]::HtmlDecode([string]$resp.responseData.translatedText)
}

function Invoke-LingvaTranslate {
  param(
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Dst,
    [Parameter(Mandatory = $false)][int]$TimeoutSec = 60
  )
  $enc = [uri]::EscapeDataString($Text)
  $url = "https://lingva.ml/api/v1/$Src/$Dst/$enc"
  $resp = Invoke-RestMethod -Uri $url -TimeoutSec $TimeoutSec
  if (-not $resp -or -not $resp.translation) {
    throw "Lingva translate failed (empty response)"
  }
  return [string]$resp.translation
}

function Invoke-Translate {
  param(
    [Parameter(Mandatory = $true)][string]$Provider,
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Dst,
    [Parameter(Mandatory = $false)][string]$Email
  )
  if ($Provider -eq "mymemory") {
    return Invoke-MyMemoryTranslate -Text $Text -Src $Src -Dst $Dst -Email $Email
  }
  if ($Provider -eq "lingva") {
    return Invoke-LingvaTranslate -Text $Text -Src $Src -Dst $Dst
  }
  throw "Unknown provider: $Provider"
}

function Translate-Text {
  param(
    [Parameter(Mandatory = $true)][string]$Text,
    [Parameter(Mandatory = $true)][string]$Provider,
    [Parameter(Mandatory = $true)][string]$Src,
    [Parameter(Mandatory = $true)][string]$Dst,
    [Parameter(Mandatory = $true)][hashtable]$Cache,
    [Parameter(Mandatory = $false)][string]$Email,
    [Parameter(Mandatory = $false)][int]$MaxLen = 450,
    [Parameter(Mandatory = $false)][int]$SleepMs = 200,
    [Parameter(Mandatory = $false)][ref]$Stats
  )
  $chunks = Split-TextForMyMemory -Text $Text -MaxLen $MaxLen
  $out = New-Object System.Text.StringBuilder

  $queue = New-Object System.Collections.Generic.Queue[string]
  foreach ($c in $chunks) { $queue.Enqueue($c) }

  while ($queue.Count -gt 0) {
    $c = $queue.Dequeue()

    if (-not (Has-Cyrillic -Text $c)) {
      [void]$out.Append($c)
      continue
    }

    if ($Cache.ContainsKey($c)) {
      if ($Stats) { $Stats.Value.cache_hits++ }
      [void]$out.Append([string]$Cache[$c])
      continue
    }

    try {
      if ($Stats) { $Stats.Value.api_calls++ }
      $tr = Invoke-Translate -Provider $Provider -Text $c -Src $Src -Dst $Dst -Email $Email
      $Cache[$c] = $tr
      [void]$out.Append($tr)
      if ($SleepMs -gt 0) { Start-Sleep -Milliseconds $SleepMs }
    }
    catch {
      # Some providers (notably Lingva) can return 404 for very long/complex URL paths.
      # Recover by splitting the chunk into smaller parts and retrying.
      if ($Provider -eq "lingva" -and $c.Length -gt 120) {
        $half = [Math]::Max(80, [int]([Math]::Floor($c.Length / 2)))
        $parts = Split-TextForMyMemory -Text $c -MaxLen $half
        foreach ($p in $parts) { $queue.Enqueue($p) }
        continue
      }
      throw
    }
  }

  return $out.ToString()
}

function Set-NodeTextSafe {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode]$TNode,
    [Parameter(Mandatory = $true)][AllowEmptyString()][string]$Text
  )
  # Word requires xml:space="preserve" when text begins/ends with space.
  if ($Text -match "^\s" -or $Text -match "\s$") {
    $null = $TNode.SetAttribute("xml:space", "preserve")
  }
  $TNode.InnerText = $Text
}

function Distribute-TextAcrossNodes {
  param(
    [Parameter(Mandatory = $true)][System.Xml.XmlNode[]]$TNodes,
    [Parameter(Mandatory = $true)][string]$Translated
  )
  if ($TNodes.Count -eq 1) {
    Set-NodeTextSafe -TNode $TNodes[0] -Text $Translated
    return
  }

  $origLens = @()
  $total = 0
  foreach ($n in $TNodes) {
    $l = ([string]$n.InnerText).Length
    $origLens += $l
    $total += $l
  }
  if ($total -le 0) {
    # Fallback: place all text into first node.
    Set-NodeTextSafe -TNode $TNodes[0] -Text $Translated
    for ($i = 1; $i -lt $TNodes.Count; $i++) { Set-NodeTextSafe -TNode $TNodes[$i] -Text "" }
    return
  }

  $tLen = $Translated.Length
  $cuts = @()
  $acc = 0
  for ($i = 0; $i -lt $TNodes.Count; $i++) {
    if ($i -eq $TNodes.Count - 1) {
      $cuts += ($tLen - $acc)
      break
    }
    $share = [int][Math]::Round($tLen * ($origLens[$i] / [double]$total))
    if ($share -lt 0) { $share = 0 }
    if ($acc + $share -gt $tLen) { $share = $tLen - $acc }
    $cuts += $share
    $acc += $share
  }

  $pos = 0
  for ($i = 0; $i -lt $TNodes.Count; $i++) {
    $len = $cuts[$i]
    if ($len -lt 0) { $len = 0 }
    if ($pos + $len -gt $tLen) { $len = $tLen - $pos }
    $seg = ""
    if ($len -gt 0) { $seg = $Translated.Substring($pos, $len) }
    Set-NodeTextSafe -TNode $TNodes[$i] -Text $seg
    $pos += $len
  }
}

# Defaults
$repoRoot = (Get-Location).Path
if (-not $CachePath) {
  $CachePath = Join-Path $repoRoot ("tmp\\translate_cache_{0}_{1}.json" -f $SourceLang, $TargetLang)
}
if (-not $ReportPath) {
  $ReportPath = Join-Path $repoRoot ("tmp\\reports\\translate_{0}_{1}.json" -f ([IO.Path]::GetFileNameWithoutExtension($OutputDocx)), $TargetLang)
}
Ensure-ParentDir -Path $ReportPath

$stats = [ordered]@{
  input = (Resolve-Path -LiteralPath $InputDocx).Path
  output = (Resolve-Path -LiteralPath (Split-Path -Parent $OutputDocx)).Path + "\\" + (Split-Path -Leaf $OutputDocx)
  provider = $Provider
  source_lang = $SourceLang
  target_lang = $TargetLang
  max_chunk_chars = $MaxChunkChars
  sleep_ms = $SleepMs
  api_calls = 0
  cache_hits = 0
  xml_files_processed = 0
  paragraphs_seen = 0
  paragraphs_translated = 0
  text_nodes_seen = 0
  text_nodes_translated = 0
}

$cache = Load-Cache -Path $CachePath

$workRoot = Join-Path (Split-Path -Parent $OutputDocx) "tmp\\work"
$workDir = Join-Path $workRoot ("docx_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
  Expand-Docx -DocxPath $InputDocx -DestinationDir $workDir

  $wordDir = Join-Path $workDir "word"
  $targets =
    Get-ChildItem -LiteralPath $wordDir -Filter "*.xml" -File |
    Where-Object {
      $_.Name -in @("document.xml", "footnotes.xml", "endnotes.xml", "comments.xml") -or
      $_.Name -like "header*.xml" -or
      $_.Name -like "footer*.xml"
    }

  foreach ($f in $targets) {
    $xml = Load-XmlDocument -Path $f.FullName
    $nsm = New-WordNamespaceManager -Xml $xml
    $paras = @($xml.SelectNodes("//w:p", $nsm))
    if (-not $paras -or $paras.Count -eq 0) { continue }

    $stats.xml_files_processed++
    for ($p = 0; $p -lt $paras.Count; $p++) {
      $stats.paragraphs_seen++
      $tNodes = @($paras[$p].SelectNodes(".//w:t", $nsm))
      if (-not $tNodes -or $tNodes.Count -eq 0) { continue }

      $stats.text_nodes_seen += $tNodes.Count
      $orig = ($tNodes | ForEach-Object { [string]$_.InnerText }) -join ""
      if (-not (Has-Cyrillic -Text $orig)) { continue }

      $translated = Translate-Text `
        -Text $orig `
        -Provider $Provider `
        -Src $SourceLang `
        -Dst $TargetLang `
        -Cache $cache `
        -Email $Email `
        -MaxLen $MaxChunkChars `
        -SleepMs $SleepMs `
        -Stats ([ref]$stats)

      if ($translated -ne $orig) {
        Distribute-TextAcrossNodes -TNodes $tNodes -Translated $translated
        $stats.paragraphs_translated++
        $stats.text_nodes_translated += $tNodes.Count
      }
    }

    Save-XmlDocumentUtf8NoBom -Xml $xml -Path $f.FullName
  }

  Save-Cache -Cache $cache -Path $CachePath
  Compress-Docx -SourceDir $workDir -OutDocxPath $OutputDocx
}
finally {
  # Report
  try { Save-Cache -Cache $cache -Path $CachePath } catch {}
  ($stats | ConvertTo-Json -Depth 5) | Set-Content -LiteralPath $ReportPath -Encoding UTF8
  if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
  }
}

Write-Host ("DONE: translated {0} -> {1} (api_calls={2}, cache_hits={3})" -f (Split-Path -Leaf $InputDocx), (Split-Path -Leaf $OutputDocx), $stats.api_calls, $stats.cache_hits)
