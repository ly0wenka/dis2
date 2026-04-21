param(
  [Parameter(Mandatory = $true)][string]$InputDocx,
  [Parameter(Mandatory = $true)][string]$OutputDocx,
  [Parameter(Mandatory = $false)][ValidateSet("Whole")][string]$Scope = "Whole",
  [Parameter(Mandatory = $false)][switch]$Fix44,
  [Parameter(Mandatory = $false)][switch]$Remove44Formulas,
  [Parameter(Mandatory = $false)][switch]$NormalizeBracketCitations,
  [Parameter(Mandatory = $false)][switch]$RemoveRefTokens,
  [Parameter(Mandatory = $false)][ValidateSet("KeepTextRemoveStrike", "DeleteText")][string]$StrikeHandling = "KeepTextRemoveStrike",
  [Parameter(Mandatory = $false)][string]$ReportPath
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

. "$PSScriptRoot\\docx_lib.ps1"

function Ensure-ParentDir {
  param([Parameter(Mandatory = $true)][string]$Path)
  $parent = Split-Path -Parent $Path
  if ($parent -and -not (Test-Path -LiteralPath $parent)) {
    New-Item -ItemType Directory -Path $parent | Out-Null
  }
}

$workRoot = Join-Path (Split-Path -Parent $OutputDocx) "tmp\\work"
$workDir = Join-Path $workRoot ("docx_" + [Guid]::NewGuid().ToString("N"))
New-Item -ItemType Directory -Path $workDir -Force | Out-Null

try {
  Expand-Docx -DocxPath $InputDocx -DestinationDir $workDir

  $docXmlPath = Join-Path $workDir "word\\document.xml"
  $xml = Load-XmlDocument -Path $docXmlPath
  $nsm = New-WordNamespaceManager -Xml $xml

  $stats = [ordered]@{
    input = (Resolve-Path -LiteralPath $InputDocx).Path
    output = (Resolve-Path -LiteralPath (Split-Path -Parent $OutputDocx)).Path + "\\" + (Split-Path -Leaf $OutputDocx)
    removed_ref_tokens = 0
    normalized_citations = 0
    bare_citations_wrapped = 0
    bare_citations_skipped = @()
    strike_runs_processed = 0
    strike_runs_deleted = 0
    strike_marks_removed = 0
    fix44_para_start = $null
    fix44_para_end = $null
    fix44_math_paras_removed = 0
    fix44_list_paras_flattened = 0
    fix44_short_paras_merged = 0
    fix44_panel_dupes_removed = 0
    fix44_detection_paras_removed = 0
    fix44_detection_text_removed = 0
    fix44_ref_paras_removed = 0
    fix44_ref_inserted = 0
    fix44_paras_deleted = 0
    fix44_paras_edited = 0
    fix44_dedup_removed = 0
  }

  # Whole document text transforms
  foreach ($t in $xml.SelectNodes("//w:t", $nsm)) {
    $orig = $t.InnerText
    $new = $orig
    if ($RemoveRefTokens) { $new = Remove-RefTokensInText $new }
    # NOTE: Per-node bracket normalization misses cases where "[" and "46; 53" live in different runs.
    # We keep this fast path for simple cases, but do the correct paragraph-level pass below.
    if ($NormalizeBracketCitations) { $new = Normalize-BracketCitationsInText $new }
    if ($new -ne $orig) {
      if ($RemoveRefTokens -and ($orig -match "REF\s+_Ref") -and ($new -notmatch "REF\s+_Ref")) { $stats.removed_ref_tokens++ }
      if ($NormalizeBracketCitations -and ($orig -match "\[[^\]]*;[^\]]*\]") -and ($new -notmatch "\[[^\]]*;[^\]]*\]")) { $stats.normalized_citations++ }
      $t.InnerText = $new
    }
  }

  function Set-ParagraphTextPlain {
    param(
      [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
      [Parameter(Mandatory = $true)][string]$Text,
      [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
    )
    $tNodes = @($Paragraph.SelectNodes(".//w:t", $Nsm))
    if ($tNodes.Count -eq 0) { return $false }
    $first = $tNodes[0]
    $first.InnerText = $Text
    for ($k = $tNodes.Count - 1; $k -ge 1; $k--) {
      [void]$tNodes[$k].ParentNode.RemoveChild($tNodes[$k])
    }
    return $true
  }

  function Rewrite-ParagraphTextKeepDrawings {
    param(
      [Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph,
      [Parameter(Mandatory = $true)][string]$Text,
      [Parameter(Mandatory = $true)][System.Xml.XmlNamespaceManager]$Nsm
    )
    $pPr = $Paragraph.SelectSingleNode("./w:pPr", $Nsm)
    $keepRuns = @()
    foreach ($r in @($Paragraph.SelectNodes("./w:r", $Nsm))) {
      if ($r.SelectSingleNode(".//w:drawing | .//w:pict | .//w:object", $Nsm)) {
        # Clone and strip any text nodes so we don't accidentally re-introduce boilerplate from runs that
        # contain both text and an inline drawing.
        $clone = $Paragraph.OwnerDocument.ImportNode($r, $true)
        foreach ($tn in @($clone.SelectNodes(".//w:t | .//w:instrText | .//w:delText", $Nsm))) {
          [void]$tn.ParentNode.RemoveChild($tn)
        }
        $keepRuns += $clone
      }
    }

    # Remove all existing children.
    foreach ($child in @($Paragraph.ChildNodes)) {
      [void]$Paragraph.RemoveChild($child)
    }

    if ($pPr) {
      [void]$Paragraph.AppendChild($Paragraph.OwnerDocument.ImportNode($pPr, $true))
    }

    if ($Text) {
      $rNew = $Paragraph.OwnerDocument.CreateElement("w", "r", $Nsm.LookupNamespace("w"))
      $tNew = $Paragraph.OwnerDocument.CreateElement("w", "t", $Nsm.LookupNamespace("w"))
      $tNew.SetAttribute("xml:space", "preserve")
      $tNew.InnerText = $Text
      [void]$rNew.AppendChild($tNew)
      [void]$Paragraph.AppendChild($rNew)
    }

    foreach ($r in $keepRuns) {
      [void]$Paragraph.AppendChild($r)
    }
    return $true
  }

  function Normalize-BracketCitationsInParagraph {
    param([Parameter(Mandatory = $true)][System.Xml.XmlNode]$Paragraph)
    $replaced = 0
    $tNodes = @($Paragraph.SelectNodes(".//w:t", $nsm))
    if ($tNodes.Count -eq 0) { return 0 }

    $inBracket = $false
    foreach ($t in $tNodes) {
      $s = $t.InnerText
      if (-not $s) { continue }
      if ($s.IndexOf("[") -lt 0 -and $s.IndexOf("]") -lt 0 -and $s.IndexOf(";") -lt 0) { continue }

      $chars = $s.ToCharArray()
      for ($i = 0; $i -lt $chars.Length; $i++) {
        if ($chars[$i] -eq "[") { $inBracket = $true; continue }
        if ($chars[$i] -eq "]") { $inBracket = $false; continue }
        if ($inBracket -and $chars[$i] -eq ";") { $chars[$i] = ","; $replaced++ }
      }
      $newS = -join $chars
      if ($newS -ne $s) { $t.InnerText = $newS }
    }
    return $replaced
  }

  function Wrap-BareCitationLists {
    param([Parameter(Mandatory = $true)][string]$Text)

    # Wrap sequences like "57; 60; 64" to "[57, 60, 64]" when outside brackets.
    # Skip sequences containing "." or "-" (mixed tokens / ranges).
    $pattern = '(?<!\[)(\b\d{1,3}(?:\s*;\s*\d{1,3}){1,})(?=\s*[\)\].,;:]|\s*$)'
    $changed = $false

    $newText = [System.Text.RegularExpressions.Regex]::Replace(
      $Text,
      $pattern,
      {
        param($m)
        # Skip if match is inside existing square brackets.
        $before = $Text.Substring(0, $m.Index)
        $open = $before.LastIndexOf('[')
        $close = $before.LastIndexOf(']')
        if ($open -gt $close) { return $m.Value }

        $list = $m.Groups[1].Value
        if ($list -match '[\\.-]') {
          if ($stats.bare_citations_skipped.Count -lt 25) {
            $stats.bare_citations_skipped += $list
          }
          return $list
        }
        $changed = $true
        $items = ($list -split ';' | ForEach-Object { $_.Trim() }) -join ', '
        return '[' + $items + ']'
      },
      [System.Text.RegularExpressions.RegexOptions]::None
    )

    if ($changed) { $stats.bare_citations_wrapped++ }
    return $newText
  }

  # Whole document: wrap bare citation lists (e.g. "57; 60; 64" -> "[57, 60, 64]")
  $allParas = $xml.SelectNodes("//w:body/w:p", $nsm)

  if ($NormalizeBracketCitations) {
    foreach ($p in $allParas) {
      $cnt = Normalize-BracketCitationsInParagraph -Paragraph $p
      if ($cnt -gt 0) { $stats.normalized_citations += $cnt }
    }
  }

  foreach ($p in $allParas) {
    $txt = Get-ParagraphText -Paragraph $p -Nsm $nsm
    if (-not $txt) { continue }
    if ($txt -notmatch '\d{1,3}\s*;\s*\d{1,3}') { continue }
    $newTxt = Wrap-BareCitationLists -Text $txt
    if ($newTxt -ne $txt) {
      [void](Set-ParagraphTextPlain -Paragraph $p -Text $newTxt -Nsm $nsm)
    }
  }

  # Whole document strikethrough
  $runs = $xml.SelectNodes("//w:r", $nsm)
  for ($i = $runs.Count - 1; $i -ge 0; $i--) {
    $r = $runs[$i]
    $rPr = $r.SelectSingleNode("./w:rPr", $nsm)
    if (-not $rPr) { continue }
    $strike = $rPr.SelectSingleNode("./w:strike | ./w:dstrike", $nsm)
    if (-not $strike) { continue }

    $stats.strike_runs_processed++
    if ($StrikeHandling -eq "DeleteText") {
      [void]$r.ParentNode.RemoveChild($r)
      $stats.strike_runs_deleted++
      continue
    }
    foreach ($node in @($rPr.SelectNodes("./w:strike | ./w:dstrike", $nsm))) {
      [void]$rPr.RemoveChild($node)
      $stats.strike_marks_removed++
    }
  }

  if ($Fix44) {
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

    function Find-ParaIndexByRegex {
      param([Parameter(Mandatory = $true)][string]$Pattern, [Parameter(Mandatory = $false)][int]$StartAt = 0)
      for ($j = $StartAt; $j -lt $paras.Count; $j++) {
        $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
        if ($txt -match $Pattern) { return $j }
      }
      return $null
    }

    function Find-ParaIndexByPrefix {
      param([Parameter(Mandatory = $true)][string]$Prefix, [Parameter(Mandatory = $false)][int]$StartAt = 0)
      for ($j = $StartAt; $j -lt $paras.Count; $j++) {
        $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
        if ($txt -and $txt.StartsWith($Prefix)) { return $j }
      }
      return $null
    }

    function Find-BodyHeadingIndexByPrefix {
      param(
        [Parameter(Mandatory = $true)][string]$Prefix,
        [Parameter(Mandatory = $false)][int]$StartAt = 0,
        [Parameter(Mandatory = $false)][int]$Lookahead = 200
      )
      $candidates = @()
      for ($j = $StartAt; $j -lt $paras.Count; $j++) {
        $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
        if ($txt -and $txt.StartsWith($Prefix)) { $candidates += $j }
      }
      if ($candidates.Count -eq 0) { return $null }
      if ($candidates.Count -eq 1) { return $candidates[0] }

      foreach ($c in $candidates) {
        $limit = [Math]::Min($paras.Count, $c + $Lookahead)
        for ($k = $c + 1; $k -lt $limit; $k++) {
          $t = Get-ParagraphText -Paragraph $paras[$k] -Nsm $nsm
          if (-not $t) { continue }
          if ($t.StartsWith("4.5")) { break }
          if (Paragraph-IsCaptionText $t) { return $c }
        }
      }
      # Fallback: last occurrence is more likely to be in the body than the TOC.
      return $candidates[$candidates.Count - 1]
    }

    # Pick the body instance of 4.4.1 (not the TOC entry) by looking ahead for figure/table captions.
    $startSearchAt = 0
    $idx441 = Find-BodyHeadingIndexByPrefix -Prefix "4.4.1" -StartAt $startSearchAt
    if ($null -eq $idx441) { throw "Could not locate heading 4.4.1 in document body." }
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    if ($null -eq $idx445) { throw "Could not locate heading 4.4.5 in document body." }
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    $stats.fix44_para_start = $idx441
    $stats.fix44_para_end = $idxAfter445

    function Insert-ParagraphAfter {
      param(
        [Parameter(Mandatory = $true)][int]$AfterIndex,
        [Parameter(Mandatory = $true)][string]$Text
      )
      $pNew = $xml.CreateElement("w", "p", $nsm.LookupNamespace("w"))
      $r = $xml.CreateElement("w", "r", $nsm.LookupNamespace("w"))
      $t = $xml.CreateElement("w", "t", $nsm.LookupNamespace("w"))
      $t.SetAttribute("xml:space", "preserve")
      $t.InnerText = $Text
      [void]$r.AppendChild($t)
      [void]$pNew.AppendChild($r)
      [void]$paras[$AfterIndex].ParentNode.InsertAfter($pNew, $paras[$AfterIndex])
      $stats.fix44_ref_inserted++
    }

    # §4.4 legend (panels + detection): ensure exactly one paragraph right after 4.4.1.
    $legendSig = "Інтерпретація панелей"
    $legendText = "Інтерпретація панелей (для всіх ілюстрацій у §4.4.1–§4.4.5): ліва панель — оптичний потік; центральна панель — карта глибини; права панель — об’єднане зображення (накладання потоку на глибину). Блок детекції DETR відображає кількість виявлених у кадрі об’єктів і використовується для прив’язки оцінених параметрів до конкретних цілей; тому службові фрази про «N об’єктів виявлено» далі не дублюються для кожного рисунка окремо."

    # Remove any existing legend duplicates in a small window under 4.4.1.
    $windowEnd = [Math]::Min($paras.Count, $idx441 + 30)
    for ($j = $windowEnd - 1; $j -ge ($idx441 + 1); $j--) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $txt) { continue }
      if ($txt.StartsWith($legendSig) -or $txt.StartsWith("У всіх наведених далі ілюстраціях розділу 4.4")) {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_dedup_removed++
      }
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

    # Insert legend once directly below 4.4.1.
    Insert-ParagraphAfter -AfterIndex $idx441 -Text $legendText
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)

    # Remove any existing reference lines (4.4.4 and 4.4.5) so we can re-insert cleanly.
    for ($j = $idxAfter445 - 1; $j -ge $idx441; $j--) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $txt) { continue }
      if ($txt -match "^Параметри\s+руху\s+обчиснюються\s+за\s+постановкою" -or $txt -match "^Параметри\s+руху\s+для\s+локальної\s+моделі") {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_ref_paras_removed++
      }
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx441 = Find-ParaIndexByPrefix -Prefix "4.4.1" -StartAt $startSearchAt
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    # Insert reference sentences under 4.4.4 and 4.4.5 headings (exactly once).
    $idx444 = Find-ParaIndexByPrefix -Prefix "4.4.4" -StartAt $idx441
    if ($null -ne $idx444) {
      Insert-ParagraphAfter -AfterIndex $idx444 -Text "Параметри руху обчиснюються за постановкою та ітераційною схемою розд. 2, формули (2.1)–(2.10)."
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    if ($null -ne $idx445) {
      Insert-ParagraphAfter -AfterIndex $idx445 -Text "Параметри руху для локальної моделі оцінювання руху наведено в розд. 2, формули (2.11)–(2.24)."
    }
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    # Rewrite/cleanup paragraphs in §4.4.1–§4.4.5 (remove boilerplate, keep concise per-figure text).
    $seen = New-Object 'System.Collections.Generic.HashSet[string]'
    for ($j = $idxAfter445 - 1; $j -ge $idx441; $j--) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $txt) { continue }

      if (Paragraph-IsHeadingText $txt -or Paragraph-IsCaptionText $txt) { continue }
      if ($txt.StartsWith($legendSig)) { continue }

      $lower = $txt.ToLowerInvariant()
      $startsPanel = $lower.StartsWith("ліва панель") -or $lower.StartsWith("центральна панель") -or $lower.StartsWith("права панель")
      if ($startsPanel) {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_paras_deleted++
        continue
      }

      # Remove generic "for comparison" paragraphs that add no new info.
      if (
        ($txt -match "^На рис\\." -and $txt -match "використовується для порівняння") -or
        ($txt -match "^На рис\\.\s*\\d+\\.\\d+\\s+наведено результат" -and $txt -match "використовується для порівняння")
      ) {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_paras_deleted++
        continue
      }

      $newTxt = $txt

      # Remove detection boilerplate (sentences/fragments).
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "Детекція об’єктів\\.?\\s*", "", [System.Text.RegularExpressions.RegexOptions]::None)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "У нижній частині[^\\.]{0,400}виявлено[^\\.]{0,200}\\.?\\s*", "", [System.Text.RegularExpressions.RegexOptions]::Singleline)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "Виявлено\\s+\\d+\\s+об.?єктів\\s+моделлю\\s+DETR\\.?\\s*", "", [System.Text.RegularExpressions.RegexOptions]::Singleline)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "Детекція\\s+\\d+\\s+об.?єктів\\s+моделлю\\s+DETR\\.?\\s*", "", [System.Text.RegularExpressions.RegexOptions]::Singleline)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "«\\s*\\d+\\s+об.?єктів\\s+виявлено\\s*»\\s*", "", [System.Text.RegularExpressions.RegexOptions]::None)
      # Remove longer DETR count sentences that vary in wording.
      $rxSingle = [System.Text.RegularExpressions.RegexOptions]::Singleline -bor [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "У\\s+нижній\\s+частині[^\\.]{0,2000}?DETR[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      # Also remove detection-count boilerplate even if the model name is missing/removed.
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "У\\s+нижній\\s+частині[^\\.]{0,2000}?виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      # Last-resort: drop any full "У нижній частині ..." sentence in §4.4 to avoid repeating service blocks.
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "У\\s+нижній\\s+частині.*?\\.\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "[,;]?\\s*зазначено,\\s*що\\s*виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}?DETR[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "[,;]?\\s*виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}?DETR[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "[,;]?\\s*зазначено,\\s*що\\s*виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "[,;]?\\s*виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "це\\s+результат\\s+роботи\\s+моделі[^\\.]{0,800}\\.?\\s*", "", $rxSingle)

      # Remove embedded panel boilerplate fragments (keep legend at top).
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "Ліва панель[^\\.]{0,600}\\.?\\s*", "", [System.Text.RegularExpressions.RegexOptions]::Singleline)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "Центральна панель[^\\.]{0,600}\\.?\\s*", "", [System.Text.RegularExpressions.RegexOptions]::Singleline)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "Права панель[^\\.]{0,600}\\.?\\s*", "", [System.Text.RegularExpressions.RegexOptions]::Singleline)
      # Also remove repeated "На лівій/центральній/правій панелі ..." sentences.
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "На\\s+лівій\\s+панелі[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "На\\s+центральній\\s+панелі[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "На\\s+правій\\s+панелі[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      # Fallback: remove panel lead-ins even if the sentence is unusually long or missing a period.
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "На\\s+лівій\\s+панелі\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "На\\s+центральній\\s+панелі\\s*", "", $rxSingle)
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "На\\s+правій\\s+панелі\\s*", "", $rxSingle)
      # Fallback: strip DETR token if it survived earlier sentence-level cleanup.
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "\\bDETR\\b", "", $rxSingle)
      # Extra-safe fallbacks (PowerShell -replace is case-insensitive).
      $newTxt = ($newTxt -replace "На лівій панелі", "") -replace "На центральній панелі", ""
      $newTxt = ($newTxt -replace "На правій панелі", "") -replace "DETR", ""
      $newTxt = ($newTxt -replace "виявлено\\s+6\\s+об.?єктів", "") -replace "виявлено 6 об'єктів", ""
      $newTxt = $newTxt -replace "виявлено 6 об’єктів", ""
      # Tidy leftovers after removing model tokens.
      $newTxt = [System.Text.RegularExpressions.Regex]::Replace($newTxt, "модел(і|лю)\\s*,\\s*", "", $rxSingle)

      # Tighten spacing.
      $newTxt = ($newTxt -replace "\s{2,}", " ").Trim()

      # If paragraph became empty after stripping boilerplate, remove it.
      if (-not $newTxt) {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_paras_deleted++
        continue
      }

      # If this is a figure-related paragraph, add a short clause referencing the legend.
      if ($newTxt.StartsWith("На рис.") -and $newTxt -notmatch "легендою") {
        $newTxt = ($newTxt.TrimEnd('.') + ". Інтерпретація панелей — за легендою на початку підрозділу.")
      }

      # Always rebuild paragraph text for stability in §4.4 (many paragraphs contain inline drawings and fragmented runs).
      if (Rewrite-ParagraphTextKeepDrawings -Paragraph $paras[$j] -Text $newTxt -Nsm $nsm) {
        if ($newTxt -ne $txt) { $stats.fix44_paras_edited++ }
      }

      # De-duplicate exact repeated paragraphs after cleanup (ignore case/whitespace).
      $key = ($newTxt.ToLowerInvariant() -replace "\s+", " ").Trim()
      if ($seen.Contains($key)) {
        [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
        $stats.fix44_dedup_removed++
      } else {
        [void]$seen.Add($key)
      }
    }

    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    if ($Remove44Formulas) {
      $toRemove = New-Object 'System.Collections.Generic.HashSet[int]'
      for ($j = $idx441; $j -lt $idxAfter445; $j++) {
        if (Paragraph-HasMath -Paragraph $paras[$j] -Nsm $nsm) {
          [void]$toRemove.Add($j)
          if ($j -gt $idx441) {
            $prevTxt = Get-ParagraphText -Paragraph $paras[$j - 1] -Nsm $nsm
            if (Paragraph-IsOnlyEquationNumberText $prevTxt) { [void]$toRemove.Add($j - 1) }
          }
          if ($j + 1 -lt $idxAfter445) {
            $nextTxt = Get-ParagraphText -Paragraph $paras[$j + 1] -Nsm $nsm
            if (Paragraph-IsOnlyEquationNumberText $nextTxt) { [void]$toRemove.Add($j + 1) }
          }
        }
      }
      for ($j = $idxAfter445 - 1; $j -ge $idx441; $j--) {
        if ($toRemove.Contains($j)) {
          [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
          $stats.fix44_math_paras_removed++
        }
      }
    }

    # Refresh
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx441 = Find-ParaIndexByPrefix -Prefix "4.4.1" -StartAt $startSearchAt
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    function Para-HasNumPr {
      param([Parameter(Mandatory = $true)][System.Xml.XmlNode]$P)
      return ($P.SelectSingleNode("./w:pPr/w:numPr", $nsm) -ne $null)
    }
    function Remove-NumPr {
      param([Parameter(Mandatory = $true)][System.Xml.XmlNode]$P)
      $numPr = $P.SelectSingleNode("./w:pPr/w:numPr", $nsm)
      if ($numPr) { [void]$numPr.ParentNode.RemoveChild($numPr) }
    }
    function Append-Runs {
      param(
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$TargetP,
        [Parameter(Mandatory = $true)][System.Xml.XmlNode]$SourceP,
        [Parameter(Mandatory = $true)][string]$Separator
      )
      $rSep = $xml.CreateElement("w", "r", $nsm.LookupNamespace("w"))
      $tSep = $xml.CreateElement("w", "t", $nsm.LookupNamespace("w"))
      $tSep.SetAttribute("xml:space", "preserve")
      $tSep.InnerText = $Separator
      [void]$rSep.AppendChild($tSep)
      [void]$TargetP.AppendChild($rSep)

      foreach ($r in @($SourceP.SelectNodes("./w:r", $nsm))) {
        [void]$TargetP.AppendChild($xml.ImportNode($r, $true))
      }
    }

    $panelSeen = @{
      left = $false
      center = $false
      right = $false
    }
    function Reset-PanelSeen { $panelSeen.left = $false; $panelSeen.center = $false; $panelSeen.right = $false }
    Reset-PanelSeen

    $j = $idx441
    while ($j -lt $idxAfter445) {
      $txt = Get-ParagraphText -Paragraph $paras[$j] -Nsm $nsm
      if (-not $txt) { $j++; continue }

      if (Paragraph-IsHeadingText $txt -or Paragraph-IsCaptionText $txt) {
        Reset-PanelSeen
        $j++; continue
      }

      $lower = $txt.ToLowerInvariant()
      $key = $null
      if ($lower -like "ліва панель*") { $key = "left" }
      elseif ($lower -like "центральна панель*") { $key = "center" }
      elseif ($lower -like "права панель*") { $key = "right" }

      if ($key) {
        if ($panelSeen[$key]) {
          [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
          $stats.fix44_panel_dupes_removed++
          $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
          $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
          $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
          if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }
          continue
        }
        $panelSeen[$key] = $true
      }

      $isListish = ($txt -like "- *") -or (Para-HasNumPr -P $paras[$j])
      if ($isListish) {
        Remove-NumPr -P $paras[$j]
        $firstT = $paras[$j].SelectSingleNode(".//w:t", $nsm)
        if ($firstT -and $firstT.InnerText -match "^\s*-\s+") {
          $firstT.InnerText = ($firstT.InnerText -replace "^\s*-\s+", "")
        }
        $stats.fix44_list_paras_flattened++
      }

      $canMerge = ($txt.Length -lt 140) -and ($j -gt $idx441)
      if ($canMerge) {
        $prevTxt = Get-ParagraphText -Paragraph $paras[$j - 1] -Nsm $nsm
        $prevBarrier = (-not $prevTxt) -or (Paragraph-IsHeadingText $prevTxt) -or (Paragraph-IsCaptionText $prevTxt)
        if (-not $prevBarrier) {
          Append-Runs -TargetP $paras[$j - 1] -SourceP $paras[$j] -Separator " "
          [void]$paras[$j].ParentNode.RemoveChild($paras[$j])
          $stats.fix44_short_paras_merged++
          $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
          $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
          $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
          if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }
          continue
        }
      }
      $j++
    }

    # Recompute bounds after paragraph merges/removals to ensure we scan the full §4.4.1–§4.4.5 range.
    $paras = $xml.SelectNodes("//w:body/w:p", $nsm)
    $idx441 = Find-ParaIndexByPrefix -Prefix "4.4.1" -StartAt $startSearchAt
    $idx445 = Find-ParaIndexByPrefix -Prefix "4.4.5" -StartAt $idx441
    $idxAfter445 = Find-ParaIndexByPrefix -Prefix "4.5" -StartAt ($idx445 + 1)
    if ($null -eq $idxAfter445) { $idxAfter445 = $paras.Count }

    # Remove remaining embedded DETR detection boilerplate inside paragraphs (when it was merged into longer text).
    for ($j = $idx441; $j -lt $idxAfter445; $j++) {
      $p = $paras[$j]
      $tNodes = @($p.SelectNodes(".//w:t", $nsm))
      if ($tNodes.Count -eq 0) { continue }

      $full = Get-ParagraphText -Paragraph $p -Nsm $nsm
      if (-not $full) { continue }
      if ($full.StartsWith($legendSig)) { continue } # keep legend

      $newFull = $full
      $newFull = $newFull -replace "Детекція об’єктів\.", ""
      $newFull = $newFull -replace "Виявлено\s+\d+\s+об.?єктів\s+моделлю\s+DETR\.\s*", ""
      $newFull = $newFull -replace "«\s*\d+\s+об.?єктів\s+виявлено\s*»\s*", ""

      $rxSingle = [System.Text.RegularExpressions.RegexOptions]::Singleline
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "Детекц[іяі].*?DETR\.\s*", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "Детекц[іяі].*?виявлено\.\s*", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "У\\s+нижній\\s+частині[^\\.]{0,2000}?виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "У\\s+нижній\\s+частині.*?\\.\\s*", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "[,;]?\\s*зазначено,\\s*що\\s*виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "[,;]?\\s*виявлено\\s+\\d+\\s+об.?єктів[^\\.]{0,2000}\\.?\\s*", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "це\\s+результат\\s+роботи\\s+моделі[^\\.]{0,800}\\.?\\s*", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "\\bDETR\\b", "", $rxSingle)
      $newFull = [System.Text.RegularExpressions.Regex]::Replace($newFull, "модел(і|лю)\\s*,\\s*", "", $rxSingle)
      $newFull = ($newFull -replace "На лівій панелі", "") -replace "На центральній панелі", ""
      $newFull = ($newFull -replace "На правій панелі", "") -replace "DETR", ""
      $newFull = ($newFull -replace "виявлено\\s+6\\s+об.?єктів", "") -replace "виявлено 6 об'єктів", ""
      $newFull = $newFull -replace "виявлено 6 об’єктів", ""
      $newFull = $newFull -replace "\s{2,}", " "
      $newFull = $newFull.Trim()

      if ($newFull -ne $full) {
        [void](Rewrite-ParagraphTextKeepDrawings -Paragraph $p -Text $newFull -Nsm $nsm)
        $stats.fix44_detection_text_removed++
      }
    }
  }

  Save-XmlDocumentUtf8NoBom -Xml $xml -Path $docXmlPath
  Compress-Docx -SourceDir $workDir -OutDocxPath $OutputDocx

  if ($ReportPath) {
    Ensure-ParentDir -Path $ReportPath
    ($stats | ConvertTo-Json -Depth 6) | Out-File -LiteralPath $ReportPath -Encoding UTF8
  }
} finally {
  if (Test-Path -LiteralPath $workDir) {
    Remove-Item -LiteralPath $workDir -Recurse -Force
  }
}
