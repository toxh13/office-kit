# word-ops.ps1 - Word-specific inspect/edit operations, dot-sourced by office.ps1.
# Keep this file ASCII (PowerShell 5.1 reads BOM-less scripts in the ANSI code page).
# Korean text only arrives through the JSON request; results are data that src/core/word.mjs turns into Korean.
# Errors are short English codes ("paragraph out of range: 9 of 8"); Invoke-WordEdit prefixes "op <n> (<name>): ".
# Every helper is prefixed with Wd so it cannot clash with the PowerPoint/Excel ops files (same scope).

# locale-proof styles (WdBuiltinStyle): key -> constant. Keys are compared without spaces/underscores/hyphens.
$WdStyleKeys = [ordered]@{
  normal = -1; heading1 = -2; heading2 = -3; heading3 = -4; heading4 = -5; heading5 = -6; heading6 = -7
  heading7 = -8; heading8 = -9; heading9 = -10; title = -63; subtitle = -75; list_bullet = -49
  list_number = -50; quote = -181; caption = -35; no_spacing = -158; toc_heading = -267
}
$WdTableStyleKeys = [ordered]@{ grid = -155; plain = -106; light_grid = -161; light_shading = -159; light_list = -160; medium_shading1 = -162 }
$WdAlign = @{ left = 0; center = 1; right = 2; justify = 3; distribute = 4 }
$WdHighlight = @{ none = 0; black = 1; blue = 2; turquoise = 3; cyan = 3; bright_green = 4; green = 4; pink = 5; red = 6; yellow = 7
  white = 8; dark_blue = 9; teal = 10; dark_green = 11; violet = 12; dark_red = 13; dark_yellow = 14; dark_gray = 15; gray = 16 }
$WdPapers = @{ a3 = @(297, 420); a4 = @(210, 297); a5 = @(148, 210); b4 = @(257, 364); b5 = @(182, 257); letter = @(215.9, 279.4); legal = @(215.9, 355.6) }

function WdKey([string]$s) { return ($s.ToLowerInvariant() -replace '[\s_\-]', '') }
function WdPt($mm) { return [single]([double]$mm * 72.0 / 25.4) }
function WdMm($pt) { return [Math]::Round([double]$pt * 25.4 / 72.0, 1) }
function WdBool($v) { if ($v) { return -1 } else { return 0 } }
function WdCut([string]$s, [int]$n) { if ($null -eq $s) { return '' }; if ($s.Length -gt $n) { return $s.Substring(0, $n) }; return $s }
# paragraph/cell text without the trailing paragraph mark and end-of-cell mark
function WdClean([string]$s) {
  if ($null -eq $s) { return '' }
  if ($s.EndsWith([string][char]7)) { $s = $s.Substring(0, $s.Length - 1) }
  if ($s.EndsWith("`r")) { $s = $s.Substring(0, $s.Length - 1) }
  return $s
}
function WdColor([string]$c) {
  $h = $c.Trim().TrimStart('#')
  if ($h -notmatch '^[0-9a-fA-F]{6}$') { throw "bad value: color: $c" }
  $r = [Convert]::ToInt32($h.Substring(0, 2), 16); $g = [Convert]::ToInt32($h.Substring(2, 2), 16); $b = [Convert]::ToInt32($h.Substring(4, 2), 16)
  return [int]($r + 256 * $g + 65536 * $b)
}
function WdColorHex($v) {
  $n = [int]$v
  if ($n -lt 0 -or $n -gt 0xFFFFFF) { return $null }
  return ('#{0:x2}{1:x2}{2:x2}' -f ($n -band 0xFF), (($n -shr 8) -band 0xFF), (($n -shr 16) -band 0xFF))
}

function WdGetStyle($doc, [string]$name, [bool]$table) {
  if (-not $name) { throw 'bad value: style' }
  $map = $(if ($table) { $WdTableStyleKeys } else { $WdStyleKeys })
  $k = WdKey $name
  foreach ($key in $map.Keys) { if ((WdKey $key) -eq $k) { return $doc.Styles.Item([int]$map[$key]) } }
  # English names of the keyed styles also work on a localized Word ('Heading 1', 'List Bullet')
  if (-not $table) { foreach ($key in $WdStyleKeys.Keys) { if ((WdKey $key) -eq $k) { return $doc.Styles.Item([int]$WdStyleKeys[$key]) } } }
  try { return $doc.Styles.Item($name) } catch { throw "style not found: $name" }
}

# builtin style local name -> key (inspect shows both)
function WdStyleKeyMap($doc) {
  $m = @{}
  foreach ($key in $WdStyleKeys.Keys) { try { $m[[string]$doc.Styles.Item([int]$WdStyleKeys[$key]).NameLocal] = $key } catch {} }
  return $m
}

function WdGetPara($doc, $n) {
  $total = [int]$doc.Paragraphs.Count
  if ($null -eq $n) { throw 'bad value: paragraph' }
  $i = [int]$n
  if ($i -lt 1 -or $i -gt $total) { throw "paragraph out of range: $i of $total" }
  return $doc.Paragraphs.Item($i)
}
function WdParaAt($doc, [int]$pos) { return $doc.Range($pos, $pos).Paragraphs.Item(1) }
# 1-based paragraph number of the paragraph that starts at (or contains) $pos
function WdParaIndexAt($doc, [int]$pos) {
  $s = [int]$doc.Range($pos, $pos).Paragraphs.Item(1).Range.Start
  if ($s -le 0) { return 1 }
  return [int]$doc.Range(0, $s).Paragraphs.Count + 1
}
function WdIsRowEnd($range) { $d = $range.Duplicate; $d.Collapse(1); return [bool]$d.IsEndOfRowMark }
# Range.Information(wdWithInTable) forces a layout pass (~30 ms); Tables.Count is cheap
function WdInTable($range) { return ([int]$range.Tables.Count -gt 0) }
# the empty paragraph of a vertically merged cell's lower part (no cell of its own; Word refuses text there)
function WdIsMergedPart($range) {
  if (-not (WdInTable $range)) { return $false }
  try { [void]$range.Cells.Item(1); return $false } catch {}
  return -not (WdIsRowEnd $range)
}
function WdCheckPara($p, $n) {
  $r = $p.Range
  if (WdInTable $r) {
    if (WdIsRowEnd $r) { throw "row end mark: $n" }
    if (WdIsMergedPart $r) { throw "merged cell: $n" }
  }
}
# a paragraph that ends in a section break (no paragraph mark: the break carries the section's page setup)
function WdIsSectionEnd($range, [string]$text) {
  if (-not $text.EndsWith([string][char]12)) { return $false }
  return ([int]$range.End -eq [int]$range.Sections.Item(1).Range.End)
}

function WdGetTable($doc, $n) {
  $total = [int]$doc.Tables.Count
  if ($null -eq $n) { throw 'bad value: table' }
  $i = [int]$n
  if ($i -lt 1 -or $i -gt $total) { throw "table out of range: $i of $total" }
  return $doc.Tables.Item($i)
}
function WdTableIndex($doc, $t) { return [int]$doc.Range(0, $t.Range.End).Tables.Count }

# case-sensitive literal search: options are set once per Find object (setting them costs ~50 ms per call)
function WdFindSetup($f) {
  $f.ClearFormatting()
  foreach ($opt in @('MatchByte', 'MatchFuzzy', 'IgnoreSpace', 'IgnorePunct', 'MatchPrefix', 'MatchSuffix', 'MatchAllWordForms', 'MatchSoundsLike', 'MatchWildcards')) {
    try { $f.$opt = ($opt -eq 'MatchByte') } catch {}
  }
}
# FindText, MatchCase, WholeWord, Wildcards, SoundsLike, AllForms, Forward, Wrap(stop), Format
function WdFindExec($f, [string]$text) { return [bool]$f.Execute($text.Replace('^', '^^'), $true, $false, $false, $false, $false, $true, 0, $false) }
# search inside $r ($r becomes the hit); same options as office.ps1 replace
function WdFindNext($r, [string]$text) { $f = $r.Find; WdFindSetup $f; return (WdFindExec $f $text) }

# result areas Word rebuilds (tables of contents/figures): finds skip them, update_fields would undo the edit
function WdTocSpans($doc) {
  $spans = New-Object System.Collections.ArrayList
  foreach ($coll in @($doc.TablesOfContents, $doc.TablesOfFigures)) {
    foreach ($x in $coll) { [void]$spans.Add(@([int]$x.Range.Start, [int]$x.Range.End)) }
  }
  return , $spans
}
# end of the span that holds $pos, or -1
function WdSpanEnd($spans, [int]$pos) { foreach ($s in $spans) { if ($pos -ge $s[0] -and $pos -lt $s[1]) { return [int]$s[1] } }; return -1 }

# next exact hit at or after $pos outside $spans, or $null (collapse-and-search-on: a fresh sub-range that
# starts inside a table finds the cell again)
function WdFindFrom($doc, [string]$t, [int]$pos, $spans) {
  $r = $doc.Range($pos, $pos)
  $f = $r.Find
  WdFindSetup $f
  for ($k = 0; $k -lt 100000; $k++) {
    if (-not (WdFindExec $f $t)) { return $null }
    $s = [int]$r.Start
    if ($s -lt $pos) { return $null }
    $skip = WdSpanEnd $spans $s
    if ($skip -ge 0) { $r.SetRange($skip, $skip); continue }
    # Word also matches look-alikes (straight/curly quotes): keep exact text only
    if ([string]$r.Text -ceq $t) { return $r }
    $r.Collapse(0)
  }
  return $null
}

# main-text characters as positions count them: a cell or row end reads as CR+BEL but takes one position
function WdPlainText([string]$s) { return $s.Replace("`r" + [string][char]7, [string][char]7) }

$WdFindCap = 5000
# hits in the main text: first one, or all (in document order). Tables of contents are skipped.
# all: positions come from one read of the main text, each checked in place; Find only re-syncs where the text and
# the positions part (fields, hidden codes). One Find per hit costs ~40 ms, this ~4 ms.
function WdFind($doc, $text, [bool]$all) {
  $t = [string]$text
  if (-not $t) { throw 'bad value: find' }
  if ($t.Length -gt 255) { throw 'find too long' }
  $spans = WdTocSpans $doc
  $hits = New-Object System.Collections.ArrayList
  if (-not $all) {
    $h = WdFindFrom $doc $t 0 $spans
    if ($h) { [void]$hits.Add($h) }
  } else {
    $L = $t.Length
    $end = [int]$doc.Content.End
    $mainText = WdPlainText ([string]$doc.Content.Text)
    $drift = 0; $last = 0
    $i = $mainText.IndexOf($t, [StringComparison]::Ordinal)
    while ($i -ge 0) {
      $next = $i + $L
      $p = $i + $drift
      if ($p -ge $last) {
        $c = $null
        if ($p + $L -le $end) { $c = $doc.Range($p, $p + $L); if ([string]$c.Text -cne $t) { $c = $null } }
        if ($c) {
          if ((WdSpanEnd $spans $p) -lt 0) { [void]$hits.Add($c) }
          $last = $p + $L
        } else {
          # re-sync: the next real hit, and its exact place in the text
          $h = WdFindFrom $doc $t $last $spans
          if (-not $h) { break }
          $j = (WdPlainText ([string]$doc.Range(0, [int]$h.Start).Text)).Length
          $drift = [int]$h.Start - $j
          [void]$hits.Add($h); $last = [int]$h.End
          $next = [Math]::Max($next, $j + $L)
        }
        if ($hits.Count -gt $WdFindCap) { throw "too many: $WdFindCap" }
      }
      $i = $mainText.IndexOf($t, $next, [StringComparison]::Ordinal)
    }
    # hits the text did not show
    while ($true) {
      $h = WdFindFrom $doc $t $last $spans
      if (-not $h) { break }
      [void]$hits.Add($h); $last = [int]$h.End
      if ($hits.Count -gt $WdFindCap) { throw "too many: $WdFindCap" }
    }
  }
  if ($hits.Count -eq 0) {
    foreach ($s in $spans) { if (([string]$doc.Range($s[0], $s[1]).Text).Contains($t)) { throw "only in toc: $t" } }
    throw "not found: $t"
  }
  return , $hits
}

# paragraph/paragraphs selector -> @{ range; from; to }
function WdSpan($doc, $op) {
  if ($null -ne $op.paragraphs) {
    $pp = @($op.paragraphs)
    if ($pp.Count -ne 2) { throw 'bad value: paragraphs' }
    $a = [int]$pp[0]; $b = [int]$pp[1]
  } elseif ($null -ne $op.paragraph) {
    $a = [int]$op.paragraph; $b = $a
  } else { return $null }
  if ($b -lt $a) { throw "bad value: paragraphs: $a > $b" }
  $pa = WdGetPara $doc $a; $pb = WdGetPara $doc $b
  return @{ range = $doc.Range($pa.Range.Start, $pb.Range.End); from = $a; to = $b }
}
# ranges an op works on: find hits or the paragraph span
function WdTargets($doc, $op, [bool]$wholeParagraphs) {
  $list = New-Object System.Collections.ArrayList
  if ($op.find) {
    # whole paragraphs: once each, however many hits it holds
    $seen = @{}
    foreach ($h in (WdFind $doc $op.find ([bool]$op.all))) {
      if (-not $wholeParagraphs) { [void]$list.Add($h); continue }
      $pr = $h.Paragraphs.Item(1).Range
      $s = [int]$pr.Start
      if (-not $seen.ContainsKey($s)) { $seen[$s] = $true; [void]$list.Add($pr) }
    }
    return , $list
  }
  $s = WdSpan $doc $op
  if ($null -eq $s) { throw 'bad value: target' }
  [void]$list.Add($s.range)
  return , $list
}

# where block content (table, TOC, page break) may not go: inside a table cell
function WdAnchor($doc, $at) {
  if ($null -eq $at) { return @{ kind = 'end' } }
  if ($at -is [string]) {
    if ($at -eq 'end') { return @{ kind = 'end' } }
    if ($at -eq 'start') { return @{ kind = 'start' } }
    throw "bad value: at: $at"
  }
  foreach ($k in @('after', 'before')) {
    $n = $at.($k + '_paragraph')
    if ($null -eq $n) { continue }
    $p = WdGetPara $doc $n
    if (WdIsMergedPart $p.Range) { throw "merged cell: $n" }
    return @{ kind = $k; para = $p }
  }
  if ($null -ne $at.after_text) {
    $hits = WdFind $doc $at.after_text $false
    return @{ kind = 'after'; para = $hits[0].Paragraphs.Item(1) }
  }
  throw 'bad value: at'
}

# new empty paragraph right after $para: split in front of its mark, so a table that follows is never entered
# (InsertParagraphAfter on the paragraph before a table puts the mark into the first cell). Returns its start.
function WdSplitAfter($doc, $para) {
  $end = [int]$para.Range.End
  $doc.Range($end - 1, $end - 1).InsertParagraphAfter()
  return $end
}

# empty paragraph in front of a table (SplitTable on the first row adds one even when the table starts the document;
# Table.Split refuses row 1)
function WdParaBeforeTable($doc, $t) {
  $start = [int]$t.Range.Start
  if ($start -gt 0) {
    $prev = WdParaAt $doc ($start - 1)
    if (-not (WdInTable $prev.Range)) { return @{ start = (WdSplitAfter $doc $prev); mode = 'after' } }
  }
  $t.Cell(1, 1).Range.Select()
  $doc.Application.Selection.SplitTable()
  return @{ start = $start; mode = 'before' }
}

# new empty paragraph at $at -> @{ start; mode('fill'|'after'|'before') }. $block: never inside a table
function WdNewPara($doc, $at, [bool]$block) {
  $a = WdAnchor $doc $at
  switch ($a.kind) {
    'end' {
      $last = $doc.Paragraphs.Last
      $lr = $last.Range
      if ([string]$lr.Text -eq "`r") {
        $prev = $last.Previous(1)
        if (-not ($block -and $prev -and (WdInTable $prev.Range))) { return @{ start = [int]$lr.Start; mode = 'fill' } }
      }
      return @{ start = (WdSplitAfter $doc $last); mode = 'after' }
    }
    'start' {
      $first = $doc.Paragraphs.First
      if (WdInTable $first.Range) { return WdParaBeforeTable $doc ($first.Range.Tables.Item(1)) }
      $first.Range.InsertParagraphBefore()
      return @{ start = 0; mode = 'before' }
    }
    'after' {
      $pr = $a.para.Range
      if ((WdInTable $pr) -and ($block -or (WdIsRowEnd $pr))) {
        $t = $pr.Tables.Item(1)
        $pos = [int]$t.Range.End
        $next = WdParaAt $doc $pos
        if (WdInTable $next.Range) { return WdParaBeforeTable $doc ($next.Range.Tables.Item(1)) }
        $next.Range.InsertParagraphBefore()
        return @{ start = $pos; mode = 'before' }
      }
      return @{ start = (WdSplitAfter $doc $a.para); mode = 'after' }
    }
    'before' {
      $pr = $a.para.Range
      if ((WdInTable $pr) -and ($block -or (WdIsRowEnd $pr))) { return WdParaBeforeTable $doc ($pr.Tables.Item(1)) }
      $start = [int]$pr.Start
      $pr.InsertParagraphBefore()
      return @{ start = $start; mode = 'before' }
    }
  }
}

# style of a new paragraph: given style, or like Enter in Word (heading -> its next style, body text -> same look).
# A split paragraph starts as a copy of its neighbour; a list item is continued like Enter does.
function WdStylePara($doc, $np, $styleName) {
  $p = WdParaAt $doc $np.start
  $restyled = $false
  if ($styleName) {
    $p.Range.Style = (WdGetStyle $doc ([string]$styleName) $false)
    $restyled = $true
  } elseif ($np.mode -ne 'fill') {
    $ref = $(if ($np.mode -eq 'before') { $p.Next(1) } else { $p.Previous(1) })
    if ($ref -and -not ((WdInTable $ref.Range) -xor (WdInTable $p.Range))) {
      $rs = $ref.Style; $ns = $rs.NextParagraphStyle
      if ([string]$ns.NameLocal -ne [string]$rs.NameLocal) { $p.Range.Style = $ns; $restyled = $true }
      # next to a picture-only paragraph: do not copy its centering/spacing
      elseif ([int]$ref.Range.InlineShapes.Count -gt 0 -and (WdClean ([string]$ref.Range.Text)).Trim('/', [char]1).Trim() -eq '') { $p.Reset() }
    }
  }
  if ($restyled) { $p.Reset(); $p.Range.Font.Reset() }
  return $restyled
}
# inherited look keeps font/size but not emphasis (bold title line -> plain new line)
function WdPlainEmphasis($range) {
  $range.HighlightColorIndex = 0
  $sf = $range.Paragraphs.Item(1).Style.Font
  $f = $range.Font
  $f.Bold = $sf.Bold; $f.Italic = $sf.Italic; $f.Underline = $sf.Underline; $f.StrikeThrough = $sf.StrikeThrough
  WdRelease $f; WdRelease $sf
}
function WdBody($para) { $r = $para.Range; [void]$r.MoveEnd(1, -1); return $r }

function WdRelease($o) { if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }

# highlight value (WdColorIndex) of a format, or $null
function WdHighlightOf($fmt) {
  if ($null -eq $fmt -or $null -eq $fmt.highlight) { return $null }
  $h = $fmt.highlight
  if ($h -is [bool]) { $h = $(if ($h) { 'yellow' } else { 'none' }) }
  $k = ([string]$h).ToLowerInvariant()
  if (-not $WdHighlight.ContainsKey($k)) { throw "bad value: highlight: $h" }
  return [int]$WdHighlight[$k]
}

# character format of one range (highlight too). Many ranges: WdFormatRanges.
function WdApplyFont($range, $fmt) {
  $n = WdApplyFontProps $range $fmt
  $hl = WdHighlightOf $fmt
  if ($null -ne $hl) { $range.HighlightColorIndex = $hl; $n++ }
  return $n
}

# After a highlight change Word does deferred work that the next call pays for once Word got idle time (a function
# call is enough) or Font is read: ~90 ms a range instead of ~2 ms. So with many ranges: font properties on every
# range first, then the highlights in one tight loop.
function WdFormatRanges($ranges, $fmt) {
  $n = 0
  foreach ($r in $ranges) { $n = WdApplyFontProps $r $fmt }
  $hl = WdHighlightOf $fmt
  if ($null -ne $hl) { foreach ($r in $ranges) { $r.HighlightColorIndex = $hl }; $n++ }
  return $n
}

function WdApplyFontProps($range, $fmt) {
  if ($null -eq $fmt) { return 0 }
  $n = 0
  $any = $false
  foreach ($key in @('bold', 'italic', 'underline', 'strike', 'color', 'size', 'font', 'superscript', 'subscript')) { if ($null -ne $fmt.$key) { $any = $true; break } }
  if (-not $any) { return 0 }
  $f = $range.Font
  try {
    if ($null -ne $fmt.bold) { $f.Bold = WdBool $fmt.bold; $n++ }
    if ($null -ne $fmt.italic) { $f.Italic = WdBool $fmt.italic; $n++ }
    if ($null -ne $fmt.underline) { $f.Underline = $(if ($fmt.underline) { 1 } else { 0 }); $n++ }
    if ($null -ne $fmt.strike) { $f.StrikeThrough = WdBool $fmt.strike; $n++ }
    if ($fmt.color) { $f.Color = WdColor $fmt.color; $n++ }
    if ($fmt.size) { $f.Size = [single]$fmt.size; $n++ }
    if ($fmt.font) { $f.Name = [string]$fmt.font; try { $f.NameFarEast = [string]$fmt.font } catch {}; $n++ }
    if ($null -ne $fmt.superscript) { $f.Superscript = WdBool $fmt.superscript; $n++ }
    if ($null -ne $fmt.subscript) { $f.Subscript = WdBool $fmt.subscript; $n++ }
  } finally { WdRelease $f }
  return $n
}

function WdApplyParaFormat($range, $op, $word) {
  $pf = $range.ParagraphFormat; $n = 0
  if ($op.align) {
    $k = ([string]$op.align).ToLowerInvariant()
    if (-not $WdAlign.ContainsKey($k)) { throw "bad value: align: $($op.align)" }
    $pf.Alignment = [int]$WdAlign[$k]; $n++
  }
  if ($op.line_spacing) { $pf.LineSpacingRule = 5; $pf.LineSpacing = [single]$word.LinesToPoints([double]$op.line_spacing); $n++ }
  if ($op.line_spacing_pt) { $pf.LineSpacingRule = 4; $pf.LineSpacing = [single]$op.line_spacing_pt; $n++ }
  if ($null -ne $op.space_before) { try { $pf.SpaceBeforeAuto = 0 } catch {}; $pf.SpaceBefore = [single]$op.space_before; $n++ }
  if ($null -ne $op.space_after) { try { $pf.SpaceAfterAuto = 0 } catch {}; $pf.SpaceAfter = [single]$op.space_after; $n++ }
  # character-unit indents (East Asian layout) win over points: clear them first
  if ($null -ne $op.first_line_indent_mm) { try { $pf.CharacterUnitFirstLineIndent = 0 } catch {}; $pf.FirstLineIndent = WdPt $op.first_line_indent_mm; $n++ }
  if ($null -ne $op.left_indent_mm) { try { $pf.CharacterUnitLeftIndent = 0 } catch {}; $pf.LeftIndent = WdPt $op.left_indent_mm; $n++ }
  if ($null -ne $op.right_indent_mm) { try { $pf.CharacterUnitRightIndent = 0 } catch {}; $pf.RightIndent = WdPt $op.right_indent_mm; $n++ }
  if ($null -ne $op.keep_with_next) { $pf.KeepWithNext = WdBool $op.keep_with_next; $n++ }
  if ($null -ne $op.page_break_before) { $pf.PageBreakBefore = WdBool $op.page_break_before; $n++ }
  return $n
}

# ---------------------------------------------------------------- inspect
function Invoke-WordInspect($doc, $word, $req) {
  $offset = 0; if ($null -ne $req.offset) { $offset = [Math]::Max(0, [int]$req.offset) }
  $limit = 200; if ($req.limit) { $limit = [Math]::Min(1000, [Math]::Max(1, [int]$req.limit)) }
  $res = @{ name = [string]$doc.Name; offset = $offset; limit = $limit }
  $res.pages = [int]$doc.ComputeStatistics(2)
  $res.words = [int]$doc.ComputeStatistics(0)
  $res.characters = [int]$doc.ComputeStatistics(3)
  $keyOf = WdStyleKeyMap $doc

  # sections: page setup + primary header/footer
  $secs = New-Object System.Collections.ArrayList
  $si = 0
  foreach ($s in $doc.Sections) {
    $si++
    if ($si -gt 20) { break }
    $ps = $s.PageSetup
    $sec = @{ index = $si; orientation = [int]$ps.Orientation; width = (WdMm $ps.PageWidth); height = (WdMm $ps.PageHeight)
      top = (WdMm $ps.TopMargin); bottom = (WdMm $ps.BottomMargin); left = (WdMm $ps.LeftMargin); right = (WdMm $ps.RightMargin)
      firstPage = [bool]$ps.DifferentFirstPageHeaderFooter }
    foreach ($kind in @('header', 'footer')) {
      $hf = $(if ($kind -eq 'header') { $s.Headers.Item(1) } else { $s.Footers.Item(1) })
      # field codes instead of stale results (a PAGE field reads '0' outside page layout)
      $hr = $hf.Range; $hr.TextRetrievalMode.IncludeFieldCodes = $true
      $sec[$kind] = WdCut (WdClean ([string]$hr.Text)) 200
      $types = @(); foreach ($f in $hf.Range.Fields) { $types += [int]$f.Type }
      $sec[$kind + 'Fields'] = $types
      if ($si -gt 1) { try { $sec[$kind + 'Linked'] = [bool]$hf.LinkToPrevious } catch {} }
      if ($sec.firstPage) {
        $fh = $(if ($kind -eq 'header') { $s.Headers.Item(2) } else { $s.Footers.Item(2) })
        $fr = $fh.Range; $fr.TextRetrievalMode.IncludeFieldCodes = $true
        $sec['first' + $kind] = WdCut (WdClean ([string]$fr.Text)) 200
      }
    }
    [void]$secs.Add($sec)
  }
  $res.sections = @($secs)
  $res.sectionCount = [int]$doc.Sections.Count

  # tables (top level, first 50)
  $tables = New-Object System.Collections.ArrayList
  $tcount = [int]$doc.Tables.Count
  for ($ti = 1; $ti -le [Math]::Min($tcount, 50); $ti++) {
    $t = $doc.Tables.Item($ti)
    $cols = 0; try { $cols = [int]$t.Columns.Count } catch {}
    $first = @()
    for ($c = 1; $c -le [Math]::Min($cols, 12); $c++) {
      $cell = $null; try { $cell = $t.Cell(1, $c) } catch {}
      if (-not $cell -or [int]$cell.RowIndex -ne 1) { break }
      $first += (WdCut (WdClean ([string]$cell.Range.Text)) 60)
    }
    $st = ''; try { $st = [string]$t.Style.NameLocal } catch {}
    [void]$tables.Add(@{ index = $ti; rows = [int]$t.Rows.Count; cols = $cols; uniform = [bool]$t.Uniform; first = $first
      style = $st; paragraph = (WdParaIndexAt $doc ([int]$t.Range.Start)) })
  }
  $res.tables = @($tables)
  $res.tableCount = $tcount

  # paragraphs: only the window [offset, offset+limit) is read (Item + Next, no walk from the start)
  $total = [int]$doc.Paragraphs.Count
  $res.paragraphTotal = $total
  $res.revisionCount = [int]$doc.Revisions.Count
  $paras = New-Object System.Collections.ArrayList
  $p = $null
  if ($offset -lt $total) { $p = $doc.Paragraphs.Item($offset + 1) }
  $i = $offset
  while ($p -and $i -lt ($offset + $limit)) {
    $i++
    $r = $p.Range
    $st = ''; try { $st = [string]$p.Style.NameLocal } catch {}
    $item = @{ i = $i; style = $st }
    if ($keyOf.ContainsKey($st)) { $item.key = $keyOf[$st] }
    $skip = $false
    if (WdInTable $r) {
      $cell = $null; try { $cell = $r.Cells.Item(1) } catch {}
      if ($cell) {
        $item.row = [int]$cell.RowIndex; $item.col = [int]$cell.ColumnIndex
        $item.table = [int]$doc.Range(0, [int]$r.End).Tables.Count
      } elseif (WdIsRowEnd $r) { $skip = $true }   # end-of-row mark: numbered but not listed
      else { $item.merged = $true; $item.table = [int]$doc.Range(0, [int]$r.End).Tables.Count }
    }
    if (-not $skip) {
      $lf = $r.ListFormat
      $lt = [int]$lf.ListType
      if ($lt -ne 0) { $item.list = [string]$lf.ListString; $item.listType = $lt }
      $txt = WdClean ([string]$r.Text)
      if (WdIsSectionEnd $r $txt) { $item.sectionBreak = $true; $txt = $txt.Substring(0, $txt.Length - 1) }
      # a paragraph deleted under track changes keeps its number until the change is accepted
      if ($res.revisionCount -gt 0 -and [int]$r.Revisions.Count -gt 0) {
        foreach ($rv in $r.Revisions) {
          if ([int]$rv.Type -eq 2 -and [int]$rv.Range.Start -le [int]$r.Start -and [int]$rv.Range.End -ge [int]$r.End) { $item.deleted = $true; break }
        }
      }
      # inline pictures read as '/' (or char 1): count them so the listing can say so
      if ($txt.Contains('/') -or $txt.Contains([string][char]1)) { $ni = [int]$r.InlineShapes.Count; if ($ni -gt 0) { $item.images = $ni } }
      if ($txt.Length -gt 200) { $item.text = $txt.Substring(0, 200); $item.cut = $txt.Length } else { $item.text = $txt }
      [void]$paras.Add($item)
    }
    $p = $p.Next(1)
  }
  $res.paragraphs = @($paras)

  # styles in use: one native Find per style that Word marks InUse (no walk over every paragraph)
  $used = New-Object System.Collections.ArrayList
  $skipName = ''; try { $skipName = [string]$doc.Styles.Item(-66).NameLocal } catch {}
  foreach ($s in $doc.Styles) {
    if (-not $s.InUse) { continue }
    $type = [int]$s.Type
    if ($type -ne 1 -and $type -ne 2) { continue }
    $name = [string]$s.NameLocal
    if ($name -eq $skipName) { continue }
    $r = $doc.Content
    $f = $r.Find
    $f.ClearFormatting()
    try { $f.Style = $s } catch { continue }
    # FindText(any), MatchCase, WholeWord, Wildcards, SoundsLike, AllForms, Forward, Wrap(stop), Format
    if ($f.Execute('', $false, $false, $false, $false, $false, $true, 0, $true)) {
      $u = @{ name = $name; character = ($type -eq 2) }
      if ($keyOf.ContainsKey($name)) { $u.key = $keyOf[$name] }
      [void]$used.Add($u)
      if ($used.Count -ge 80) { break }
    }
  }
  $res.styles = @($used)

  # comments (listed on the first page of paragraphs only)
  $comments = New-Object System.Collections.ArrayList
  $ci = 0
  $res.commentsListed = ($offset -eq 0)
  foreach ($c in $(if ($offset -eq 0) { $doc.Comments } else { @() })) {
    $ci++
    if ($ci -gt 100) { break }
    $item = @{ index = $ci; author = [string]$c.Author; text = (WdCut (WdClean ([string]$c.Range.Text)) 200); scope = (WdCut ([string]$c.Scope.Text) 60) }
    try { $item.done = [bool]$c.Done } catch {}
    [void]$comments.Add($item)
  }
  $res.comments = @($comments)
  $res.commentCount = [int]$doc.Comments.Count

  # track changes
  $res.trackRevisions = [bool]$doc.TrackRevisions
  $res.revisionCount = [int]$doc.Revisions.Count
  $revs = New-Object System.Collections.ArrayList
  if ($res.revisionCount -gt 0) {
    $ri = 0
    foreach ($rv in $doc.Revisions) {
      $ri++
      if ($ri -gt 20) { break }
      [void]$revs.Add(@{ type = [int]$rv.Type; author = [string]$rv.Author; text = (WdCut ([string]$rv.Range.Text) 60) })
    }
  }
  $res.revisions = @($revs)

  $res.counts = @{ toc = [int]$doc.TablesOfContents.Count; inlineShapes = [int]$doc.InlineShapes.Count; shapes = [int]$doc.Shapes.Count
    contentControls = [int]$doc.ContentControls.Count; bookmarks = [int]$doc.Bookmarks.Count; footnotes = [int]$doc.Footnotes.Count
    endnotes = [int]$doc.Endnotes.Count; hyperlinks = [int]$doc.Hyperlinks.Count; fields = [int]$doc.Fields.Count }
  return $res
}

# ---------------------------------------------------------------- edit ops
function WdOpInsertParagraph($doc, $word, $op) {
  $np = WdNewPara $doc $op.at $false
  $restyled = WdStylePara $doc $np $op.style
  $body = WdBody (WdParaAt $doc $np.start)
  $body.Text = [string]$op.text
  if (-not $restyled) { WdPlainEmphasis $body }
  [void](WdApplyFont $body $op.format)
  $p = WdParaAt $doc $np.start
  return @{ paragraph = (WdParaIndexAt $doc $np.start); style = [string]$p.Style.NameLocal }
}

function WdOpSetParagraphText($doc, $word, $op) {
  $p = WdGetPara $doc $op.paragraph
  WdCheckPara $p $op.paragraph
  $body = WdBody $p
  $old = [string]$body.Text
  $body.Text = [string]$op.text
  return @{ paragraph = [int]$op.paragraph; old = (WdCut $old 60) }
}

# delete one paragraph -> 'deleted' | 'section' (deleted, its section break kept) | 'tracked' | 'cleared:<why>'.
# Word keeps some paragraph marks (last of the document or of a cell, before a table) and a section break carries
# its section's page setup: those paragraphs merge into the previous one, which gets its own look back.
function WdDeletePara($doc, $p, [int]$index) {
  WdCheckPara $p $index
  $r = $p.Range
  $inT = WdInTable $r
  $txt = [string]$r.Text
  $isLast = ([int]$r.End -ge [int]$doc.Content.End)
  $cellEnd = $txt.EndsWith([string][char]7)
  $secEnd = (-not $inT) -and (WdIsSectionEnd $r $txt)
  $beforeTable = $false
  if (-not $inT -and -not $isLast) { $nx = $p.Next(1); if ($nx -and (WdInTable $nx.Range)) { $beforeTable = $true } }
  $simple = -not ($isLast -or $cellEnd -or $secEnd -or $beforeTable)
  if ([bool]$doc.TrackRevisions) {
    # tracked: the paragraph stays (marked deleted) until the change is accepted; never mark a kept mark or a break
    if ($simple) { $r.Delete() | Out-Null; return 'tracked' }
    $b = WdBody $p; if ([int]$b.End -gt [int]$b.Start) { $b.Delete() | Out-Null }
    return 'tracked-text'
  }
  if ($simple) {
    $n0 = [int]$doc.Paragraphs.Count
    $r.Delete() | Out-Null
    if ([int]$doc.Paragraphs.Count -lt $n0) { return 'deleted' }
    return 'cleared:kept'
  }
  $prev = $p.Previous(1)
  $ok = $false
  if ($prev) {
    $pr = $prev.Range
    $pt = [string]$pr.Text
    if ($cellEnd) { $ok = ((WdInTable $pr) -and -not $pt.EndsWith([string][char]7) -and -not (WdIsRowEnd $pr)) }
    else { $ok = (-not (WdInTable $pr)) -and -not (WdIsSectionEnd $pr $pt) }
  }
  if (-not $ok) {
    $b = WdBody $p
    if ([int]$b.End -gt [int]$b.Start) { $b.Delete() | Out-Null }
    $why = $(if ($cellEnd) { 'cell' } elseif ($secEnd) { 'section' } elseif ($beforeTable) { 'table' } else { 'last' })
    return "cleared:$why"
  }
  # remove the previous paragraph mark + this text, then give the merged paragraph back its own look
  $style = $prev.Style
  $fmt = $prev.Format.Duplicate
  $lf = $pr.ListFormat
  $list = $null
  if ([int]$lf.ListType -ne 0) { $list = @{ tpl = $lf.ListTemplate; level = [int]$lf.ListLevelNumber } }
  $ps = [int]$pr.Start
  $doc.Range([int]$pr.End - 1, [int]$r.End - 1).Delete() | Out-Null
  $m = WdParaAt $doc $ps
  $m.Style = $style
  $m.Format = $fmt
  if ($list -and [int]$m.Range.ListFormat.ListType -eq 0) {
    # ListTemplate, ContinuePreviousList, ApplyTo(wdListApplyToSelection 2), DefaultListBehavior(Word 10), ApplyLevel
    try { $m.Range.ListFormat.ApplyListTemplateWithLevel($list.tpl, $true, 2, 2, $list.level) } catch {}
  }
  if ($secEnd) { return 'section' }
  return 'deleted'
}

function WdOpDeleteParagraph($doc, $word, $op) {
  $deleted = @(); $cleared = @(); $tracked = @(); $trackedText = @(); $sections = @()
  $idx = @()
  if ($op.find) {
    $starts = New-Object System.Collections.ArrayList
    foreach ($h in (WdFind $doc $op.find ([bool]$op.all))) {
      $s = [int]$h.Paragraphs.Item(1).Range.Start
      if (-not $starts.Contains($s)) { [void]$starts.Add($s) }
    }
    $idx = @($starts | ForEach-Object { WdParaIndexAt $doc $_ })
  } else {
    [void](WdGetPara $doc $op.paragraph)
    $idx = @([int]$op.paragraph)
  }
  # from the end, so the numbers still to do stay valid
  for ($k = $idx.Count - 1; $k -ge 0; $k--) {
    $n = [int]$idx[$k]
    $how = WdDeletePara $doc ($doc.Paragraphs.Item($n)) $n
    switch -Wildcard ($how) {
      'deleted' { $deleted += $n }
      'section' { $deleted += $n; $sections += $n }
      'tracked' { $tracked += $n }
      'tracked-text' { $trackedText += $n }
      'cleared:*' { $cleared += @{ paragraph = $n; why = $how.Substring(8) } }
    }
  }
  [array]::Sort($deleted); [array]::Sort($tracked); [array]::Sort($trackedText); [array]::Sort($sections)
  [array]::Reverse($cleared)
  return @{ deleted = $deleted; cleared = $cleared; tracked = $tracked; trackedText = $trackedText; sectionKept = $sections }
}

function WdOpApplyStyle($doc, $word, $op) {
  $style = WdGetStyle $doc ([string]$op.style) $false
  $isChar = ([int]$style.Type -eq 2)
  $count = 0
  foreach ($r in (WdTargets $doc $op (-not $isChar))) { $r.Style = $style; $count++ }
  $res = @{ style = [string]$style.NameLocal; count = $count; character = $isChar }
  if (-not $op.find) { $res.from = [int]$(if ($null -ne $op.paragraphs) { @($op.paragraphs)[0] } else { $op.paragraph }); $res.to = [int]$(if ($null -ne $op.paragraphs) { @($op.paragraphs)[1] } else { $op.paragraph }) }
  return $res
}

function WdOpFormatText($doc, $word, $op) {
  # paragraph targets keep their marks: list numbers follow the text, the same for every paragraph
  $targets = WdTargets $doc $op $false
  if ((WdFormatRanges $targets $op) -eq 0) { throw 'bad value: no format given' }
  return @{ count = $targets.Count }
}

function WdOpFormatParagraph($doc, $word, $op) {
  $count = 0
  foreach ($r in (WdTargets $doc $op $true)) {
    if ((WdApplyParaFormat $r $op $word) -eq 0) { throw 'bad value: no format given' }
    $count += [int]$r.Paragraphs.Count
  }
  return @{ count = $count }
}

function WdOpSetList($doc, $word, $op) {
  $s = WdSpan $doc $op
  if ($null -eq $s) { throw 'bad value: paragraphs' }
  $r = $s.range
  $type = ([string]$op.type).ToLowerInvariant()
  $texts = $null
  switch ($type) {
    'none' { $r.ListFormat.RemoveNumbers(1) }
    { $_ -in @('bullet', 'number') } {
      # Word turns a typed '1. ' / '- ' at the start into the automatic number: report what it took away
      $texts = @(); $p = WdGetPara $doc $s.from
      for ($i = $s.from; $i -le $s.to -and $p; $i++) { $texts += [string](WdBody $p).Text; $p = $p.Next(1) }
      $gallery = $(if ($type -eq 'bullet') { 1 } else { 2 })
      $tpl = $word.ListGalleries.Item($gallery).ListTemplates.Item(1)
      # ListTemplate, ContinuePreviousList(false = restart), ApplyTo(wdListApplyToSelection 2), DefaultListBehavior(Word 10)
      $r.ListFormat.ApplyListTemplateWithLevel($tpl, $false, 2, 2)
    }
    default { throw "bad value: type: $($op.type)" }
  }
  $res = @{ from = $s.from; to = $s.to; type = $type }
  if ($null -ne $texts) {
    # Word reads a typed '1. ' / '- ' as numbering: it drops it from the text and may start lists of its own
    # ('- ' -> bullets, 'A.' letters). With that text gone, numbering the span once more gives one list.
    $redo = $false
    $p = WdGetPara $doc $s.from
    for ($k = 0; $k -lt $texts.Count -and $p; $k++) {
      if ([string](WdBody $p).Text -ne $texts[$k]) { $redo = $true; break }
      $p = $p.Next(1)
    }
    if ($redo) {
      $r.ListFormat.RemoveNumbers(1)
      $r.ListFormat.ApplyListTemplateWithLevel($tpl, $false, 2, 2)
    }
    $removed = New-Object System.Collections.ArrayList
    $p = WdGetPara $doc $s.from
    for ($k = 0; $k -lt $texts.Count -and $p; $k++) {
      $b = $texts[$k]; $a = [string](WdBody $p).Text
      if ($a -ne $b -and $a.Length -lt $b.Length -and $b.EndsWith($a)) { [void]$removed.Add(@{ paragraph = $s.from + $k; removed = $b.Substring(0, $b.Length - $a.Length) }) }
      $p = $p.Next(1)
    }
    if ($removed.Count) { $res.removed = @($removed) }
  }
  return $res
}

function WdOpInsertTable($doc, $word, $op) {
  $rows = @($op.rows)
  $nr = $rows.Count; $nc = 0
  foreach ($row in $rows) { $nc = [Math]::Max($nc, @($row).Count) }
  if ($nr -lt 1 -or $nc -lt 1) { throw 'bad value: rows' }
  $np = WdNewPara $doc $op.at $true
  $p = WdParaAt $doc $np.start
  # a table right after another table would merge with it: keep one empty paragraph between
  $prev = $p.Previous(1)
  if ($prev -and (WdInTable $prev.Range)) { $p = WdParaAt $doc (WdSplitAfter $doc $p) }
  $start = [int]$p.Range.Start
  $p.Range.Style = $doc.Styles.Item(-1); $p.Reset(); $p.Range.Font.Reset()
  $fit = 2
  if ($op.autofit) { switch ([string]$op.autofit) { 'fixed' { $fit = 0 } 'content' { $fit = 1 } 'window' { $fit = 2 } default { throw "bad value: autofit: $($op.autofit)" } } }
  $r = $p.Range; $r.Collapse(1)
  # Range, rows, cols, wdWord9TableBehavior (grid borders), AutoFitBehavior
  $t = $doc.Tables.Add($r, $nr, $nc, 1, $fit)
  if ($op.style) { $t.Style = [string](WdGetStyle $doc ([string]$op.style) $true).NameLocal }
  for ($i = 0; $i -lt $nr; $i++) {
    $vals = @($rows[$i])
    for ($j = 0; $j -lt $vals.Count; $j++) {
      if ($null -ne $vals[$j] -and [string]$vals[$j] -ne '') { $t.Cell($i + 1, $j + 1).Range.Text = [string]$vals[$j] }
    }
  }
  if ($null -eq $op.header_bold -or $op.header_bold) { $t.Rows.Item(1).Range.Font.Bold = -1; $t.Rows.Item(1).HeadingFormat = -1 }
  # Word keeps the empty insertion paragraph after the table: drop it unless it ends the document or separates tables
  $after = WdParaAt $doc ([int]$t.Range.End)
  if ([string]$after.Range.Text -eq "`r") {
    $nx = $after.Next(1)
    if ($nx -and -not (WdInTable $nx.Range)) { [void]$after.Range.Delete() }
  }
  return @{ table = (WdTableIndex $doc $t); rows = $nr; cols = $nc; paragraph = (WdParaIndexAt $doc ([int]$t.Range.Start))
    after = (WdParaIndexAt $doc ([int]$t.Range.End)) }
}

function WdOpTableSetCell($doc, $word, $op) {
  $t = WdGetTable $doc $op.table
  $cell = $null
  # Cell() clamps a too-large row to the last one: check the indices it really returned
  try { $cell = $t.Cell([int]$op.row, [int]$op.col) } catch {}
  if (-not $cell -or [int]$cell.RowIndex -ne [int]$op.row -or [int]$cell.ColumnIndex -ne [int]$op.col) {
    throw "cell out of range: $($op.row),$($op.col) of $($t.Rows.Count)x$($t.Columns.Count)"
  }
  $old = WdClean ([string]$cell.Range.Text)
  $cell.Range.Text = [string]$op.text
  return @{ table = [int]$op.table; row = [int]$op.row; col = [int]$op.col; old = (WdCut $old 60) }
}

$WdW = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
# cell layout from the table's XML (COM has no row span): one list per row of @{ idx(= Cell() column); grid; vm; down }.
# vm: '' | 'restart' | 'cont' (merged into the cell above: no Cell() of its own); down: the cell goes on into the next row
function WdTableGrid($t) {
  $x = New-Object System.Xml.XmlDocument
  $x.LoadXml([string]$t.Range.WordOpenXML)
  $ns = New-Object System.Xml.XmlNamespaceManager($x.NameTable)
  $ns.AddNamespace('w', $WdW)
  $rows = New-Object System.Collections.ArrayList
  $tbl = $x.SelectSingleNode('//w:body/w:tbl', $ns)
  if (-not $tbl) { return , $rows }
  foreach ($tr in $tbl.SelectNodes('w:tr | w:sdt/w:sdtContent/w:tr | w:customXml/w:tr', $ns)) {
    $g = 0
    $gb = $tr.SelectSingleNode('w:trPr/w:gridBefore', $ns); if ($gb) { $g = [int]$gb.GetAttribute('val', $WdW) }
    $cells = New-Object System.Collections.ArrayList
    $k = 0
    foreach ($tc in $tr.SelectNodes('w:tc | w:sdt/w:sdtContent/w:tc | w:customXml/w:tc', $ns)) {
      $k++
      $span = 1; $gs = $tc.SelectSingleNode('w:tcPr/w:gridSpan', $ns); if ($gs) { $span = [int]$gs.GetAttribute('val', $WdW) }
      $vm = ''; $v = $tc.SelectSingleNode('w:tcPr/w:vMerge', $ns)
      if ($v) { $vm = $(if ($v.GetAttribute('val', $WdW) -eq 'restart') { 'restart' } else { 'cont' }) }
      [void]$cells.Add(@{ idx = $k; grid = $g; vm = $vm; down = $false })
      $g += $span
    }
    [void]$rows.Add($cells)
  }
  for ($r = 0; $r -lt $rows.Count - 1; $r++) {
    foreach ($c in $rows[$r]) {
      if ($c.vm -eq '') { continue }
      foreach ($d in $rows[$r + 1]) { if ($d.vm -eq 'cont' -and $d.grid -eq $c.grid) { $c.down = $true; break } }
    }
  }
  return , $rows
}

# write values into a row by column (null = leave); a cell merged into the one above takes no value
function WdFillRow($t, [int]$row, $vals) {
  for ($j = 0; $j -lt $vals.Count; $j++) {
    if ($null -eq $vals[$j]) { continue }
    $cell = $null; try { $cell = $t.Cell($row, $j + 1) } catch {}
    if (-not $cell -or [int]$cell.RowIndex -ne $row -or [int]$cell.ColumnIndex -ne ($j + 1)) {
      $have = @((WdTableGrid $t)[$row - 1]).Count
      if ($j + 1 -gt $have) { throw "row cells: $($vals.Count) of $have" }
      throw "merged cell value: $row,$($j + 1)"
    }
    $cell.Range.Text = [string]$vals[$j]
  }
}

function WdOpTableAddRow($doc, $word, $op) {
  $t = WdGetTable $doc $op.table
  $nrows = [int]$t.Rows.Count
  $after = $(if ($null -ne $op.after_row) { [int]$op.after_row } else { $nrows })
  if ($after -lt 0 -or $after -gt $nrows) { throw "row out of range: $after of $nrows" }
  $vals = @(); if ($null -ne $op.values) { $vals = @($op.values) }
  $ncols = 0; try { $ncols = [int]$t.Columns.Count } catch {}
  if ($ncols -gt 0 -and $vals.Count -gt $ncols) { throw "too many values: $($vals.Count) of $ncols" }
  $new = $after + 1
  if ($after -eq $nrows) { [void]$t.Rows.Add() }
  else {
    $next = $null; try { $next = $t.Rows.Item($new) } catch {}
    if ($next) { [void]$t.Rows.Add($next) }
    else {
      # vertically merged cells (no Rows.Item): above the next row through one of its own cells - a merge running across
      # the gap takes the new row in, as in Word; failing that, below a cell of this row that ends here
      $grid = WdTableGrid $t
      if ($grid.Count -ne $nrows) { throw "merged rows: $new" }
      $sel = $null
      foreach ($c in $grid[$after]) { if ($c.vm -ne 'cont') { $sel = @($new, $c.idx, 'above'); break } }
      if (-not $sel -and $after -ge 1) { foreach ($c in $grid[$after - 1]) { if ($c.vm -ne 'cont' -and -not $c.down) { $sel = @($after, $c.idx, 'below'); break } } }
      if (-not $sel) { throw "merged rows: $new" }
      $t.Cell([int]$sel[0], [int]$sel[1]).Select()
      if ($sel[2] -eq 'above') { $word.Selection.InsertRowsAbove(1) } else { $word.Selection.InsertRowsBelow(1) }
      # wdStartOfRangeRowNumber: the selection is the new row
      $got = [int]$word.Selection.Information(13)
      if ($got -ne $new) { throw "inserted at: $got" }
    }
  }
  WdFillRow $t $new $vals
  $res = @{ table = [int]$op.table; row = $new; rows = [int]$t.Rows.Count }
  if ([bool]$doc.TrackRevisions) { $res.tracked = $true }
  return $res
}

function WdOpTableDeleteRow($doc, $word, $op) {
  $t = WdGetTable $doc $op.table
  $nrows = [int]$t.Rows.Count
  $row = [int]$op.row
  if ($row -lt 1 -or $row -gt $nrows) { throw "row out of range: $row of $nrows" }
  $tracked = [bool]$doc.TrackRevisions
  $res = @{ table = [int]$op.table; row = $row }
  if ($tracked) { $res.tracked = $true }
  if ($nrows -eq 1) { $t.Delete(); $res.rows = 0; $res.tableDeleted = $true; return $res }
  $one = $null; try { $one = $t.Rows.Item($row) } catch {}
  if ($one) { $one.Delete() }
  else {
    # vertically merged cells: delete through a cell that spans only this row (a merged cell would take all its rows).
    # A merged cell that starts here leaves its text to the part that stays (Word would drop it).
    $grid = WdTableGrid $t
    if ($grid.Count -ne $nrows) { throw "merged rows: $row" }
    $cells = $grid[$row - 1]
    $via = $null
    foreach ($c in $cells) { if ($c.vm -ne 'cont' -and -not $c.down) { $via = $c; break } }
    if (-not $via) { throw "merged rows: $row" }
    $keep = @()
    if (-not $tracked) {
      foreach ($c in $cells) {
        if (-not $c.down -or $c.vm -eq 'cont') { continue }
        $txt = WdClean ([string]$t.Cell($row, $c.idx).Range.Text)
        if (-not $txt) { continue }
        foreach ($d in $grid[$row]) { if ($d.vm -eq 'cont' -and $d.grid -eq $c.grid) { $keep += @{ idx = $d.idx; text = $txt }; break } }
      }
    }
    $t.Cell($row, $via.idx).Delete(2)
    if (-not $tracked -and [int]$t.Rows.Count -ne $nrows - 1) { throw "merged rows: $row" }
    $kept = @()
    foreach ($k in $keep) {
      $cell = $null; try { $cell = $t.Cell($row, $k.idx) } catch {}
      if ($cell -and -not (WdClean ([string]$cell.Range.Text))) { $cell.Range.Text = $k.text; $kept += (WdCut $k.text 30) }
    }
    if ($kept.Count) { $res.kept = $kept }
  }
  $res.rows = [int]$t.Rows.Count
  return $res
}

function WdOpInsertImage($doc, $word, $op) {
  $path = [string]$op.image_path
  if (-not $path -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { throw "image not found: $path" }
  $np = WdNewPara $doc $op.at $false
  [void](WdStylePara $doc $np $null)
  $p = WdParaAt $doc $np.start
  $r = $p.Range; $r.Collapse(1)
  # FileName, LinkToFile, SaveWithDocument, Range
  $shape = $doc.InlineShapes.AddPicture($path, $false, $true, $r)
  $w0 = [double]$shape.Width; $h0 = [double]$shape.Height
  $w = $w0; $h = $h0
  if ($op.width_mm -and $op.height_mm) { $w = [double](WdPt $op.width_mm); $h = [double](WdPt $op.height_mm) }
  elseif ($op.width_mm) { $w = [double](WdPt $op.width_mm); $h = $h0 * $w / $w0 }
  elseif ($op.height_mm) { $h = [double](WdPt $op.height_mm); $w = $w0 * $h / $h0 }
  else {
    # no size given: shrink to the text width of the page
    $ps = $p.Range.Sections.Item(1).PageSetup
    $max = [double]$ps.PageWidth - [double]$ps.LeftMargin - [double]$ps.RightMargin
    if ($w0 -gt $max -and $max -gt 0) { $w = $max; $h = $h0 * $max / $w0 }
  }
  $shape.LockAspectRatio = 0
  $shape.Width = [single]$w; $shape.Height = [single]$h
  if (-not ($op.width_mm -and $op.height_mm)) { $shape.LockAspectRatio = -1 }
  $p = WdParaAt $doc $np.start
  $pf = $p.Range.ParagraphFormat
  $alignKey = $(if ($op.align) { ([string]$op.align).ToLowerInvariant() } else { 'center' })
  if (-not $WdAlign.ContainsKey($alignKey)) { throw "bad value: align: $($op.align)" }
  $pf.Alignment = [int]$WdAlign[$alignKey]
  # exact line spacing would clip the picture
  $pf.LineSpacingRule = 0
  try { $pf.CharacterUnitFirstLineIndent = 0 } catch {}
  $pf.FirstLineIndent = 0
  $res = @{ paragraph = (WdParaIndexAt $doc $np.start); width = (WdMm $shape.Width); height = (WdMm $shape.Height) }
  if ($op.caption) {
    $label = $(if ($op.caption_label) { [string]$op.caption_label } else { 'Figure' })
    $capStart = WdSplitAfter $doc $p
    $cp = WdParaAt $doc $capStart
    $cp.Range.Style = $doc.Styles.Item(-35)
    $cp.Reset(); $cp.Range.Font.Reset()
    $cp.Range.ParagraphFormat.Alignment = [int]$WdAlign[$alignKey]
    $body = WdBody $cp
    $body.Text = "$label {seq}. " + [string]$op.caption
    $hit = (WdParaAt $doc $capStart).Range
    if (WdFindNext $hit '{seq}') { [void]$doc.Fields.Add($hit, -1, ('SEQ ' + ($label -replace '\s', '_') + ' \* ARABIC'), $false) }
    $res.caption = (WdParaIndexAt $doc $capStart)
  }
  return $res
}

function WdOpInsertPageBreak($doc, $word, $op) {
  $np = WdNewPara $doc $op.at $true
  $body = WdBody (WdParaAt $doc $np.start)
  $body.Text = [string][char]12
  $i = WdParaIndexAt $doc $np.start
  return @{ paragraph = $i }
}

function WdOpHeaderFooter($doc, $word, $op, [bool]$footer) {
  $text = $(if ($null -ne $op.text) { [string]$op.text } else { '' })
  if ($op.page_numbers -and $text -notmatch '\{page\}') { $text = $(if ($text) { "$text {page}" } else { '{page}' }) }
  $secs = @()
  if ($null -ne $op.section) {
    $total = [int]$doc.Sections.Count
    $si = [int]$op.section
    if ($si -lt 1 -or $si -gt $total) { throw "section out of range: $si of $total" }
    $secs = @($doc.Sections.Item($si))
  } else { foreach ($s in $doc.Sections) { $secs += $s } }
  $done = 0; $k = 0
  foreach ($s in $secs) {
    $k++
    $hf = $(if ($footer) { $s.Footers.Item(1) } else { $s.Headers.Item(1) })
    if ($null -ne $op.section) { if ([int]$op.section -gt 1) { $hf.LinkToPrevious = $false } }
    elseif ($k -gt 1 -and $hf.LinkToPrevious) { continue }
    $hf.Range.Text = $text
    foreach ($ph in @(@('{page}', 'PAGE'), @('{pages}', 'NUMPAGES'))) {
      for ($guard = 0; $guard -lt 20; $guard++) {
        $hit = $hf.Range
        if (-not (WdFindNext $hit $ph[0])) { break }
        [void]$doc.Fields.Add($hit, -1, $ph[1], $false)
      }
    }
    [void]$hf.Range.Fields.Update()
    $alignKey = $null
    if ($op.align) { $alignKey = ([string]$op.align).ToLowerInvariant() } elseif ($footer -and $op.page_numbers) { $alignKey = 'center' }
    if ($alignKey) {
      if (-not $WdAlign.ContainsKey($alignKey)) { throw "bad value: align: $($op.align)" }
      $hf.Range.ParagraphFormat.Alignment = [int]$WdAlign[$alignKey]
    }
    $done++
  }
  return @{ sections = $done; total = [int]$doc.Sections.Count; text = $text }
}

function WdOpPageSetup($doc, $word, $op) {
  $secs = @()
  if ($null -ne $op.section) {
    $total = [int]$doc.Sections.Count
    $si = [int]$op.section
    if ($si -lt 1 -or $si -gt $total) { throw "section out of range: $si of $total" }
    $secs = @($doc.Sections.Item($si))
  } else { foreach ($s in $doc.Sections) { $secs += $s } }
  $paper = $null
  if ($op.paper) {
    $paper = $WdPapers[([string]$op.paper).ToLowerInvariant()]
    if (-not $paper) { throw "bad value: paper: $($op.paper)" }
  }
  $orient = $null
  if ($op.orientation) {
    switch (([string]$op.orientation).ToLowerInvariant()) { 'portrait' { $orient = 0 } 'landscape' { $orient = 1 } default { throw "bad value: orientation: $($op.orientation)" } }
  }
  foreach ($s in $secs) {
    $ps = $s.PageSetup
    # Orientation swaps width/height by itself; paper sizes are then set for that orientation
    if ($null -ne $orient -and [int]$ps.Orientation -ne $orient) { $ps.Orientation = $orient }
    if ($paper) {
      $land = ([int]$ps.Orientation -eq 1)
      $w = $paper[0]; $h = $paper[1]
      if ($land) { $w = $paper[1]; $h = $paper[0] }
      $ps.PageWidth = WdPt $w; $ps.PageHeight = WdPt $h
    }
    $m = $op.margins_mm
    if ($m) {
      if ($null -ne $m.top) { $ps.TopMargin = WdPt $m.top }
      if ($null -ne $m.bottom) { $ps.BottomMargin = WdPt $m.bottom }
      if ($null -ne $m.left) { $ps.LeftMargin = WdPt $m.left }
      if ($null -ne $m.right) { $ps.RightMargin = WdPt $m.right }
    }
  }
  $ps = $secs[0].PageSetup
  return @{ sections = $secs.Count; orientation = [int]$ps.Orientation; width = (WdMm $ps.PageWidth); height = (WdMm $ps.PageHeight)
    top = (WdMm $ps.TopMargin); bottom = (WdMm $ps.BottomMargin); left = (WdMm $ps.LeftMargin); right = (WdMm $ps.RightMargin) }
}

function WdOpAddComment($doc, $word, $op) {
  $hits = WdFind $doc $op.find ([bool]$op.all)
  for ($k = $hits.Count - 1; $k -ge 0; $k--) {
    $c = $doc.Comments.Add($hits[$k], [string]$op.text)
    if ($op.author) { try { $c.Author = [string]$op.author; $c.Initial = ([string]$op.author).Substring(0, 1) } catch {} }
  }
  return @{ count = $hits.Count; total = [int]$doc.Comments.Count }
}

function WdOpDeleteComments($doc, $word, $op) {
  $total = [int]$doc.Comments.Count
  if ($null -ne $op.index) {
    $i = [int]$op.index
    if ($i -lt 1 -or $i -gt $total) { throw "comment out of range: $i of $total" }
    $doc.Comments.Item($i).Delete()
    return @{ deleted = 1; total = [int]$doc.Comments.Count }
  }
  if (-not $op.all) { throw 'bad value: all or index' }
  if ($total -gt 0) { $doc.DeleteAllComments() }
  return @{ deleted = $total; total = [int]$doc.Comments.Count }
}

function WdStoryRevisions($doc) {
  $n = 0
  foreach ($story in $doc.StoryRanges) { $r = $story; while ($r) { $n += [int]$r.Revisions.Count; $r = $r.NextStoryRange } }
  return $n
}

function WdOpInsertToc($doc, $word, $op) {
  $levels = 3
  if ($op.levels) { $levels = [int]$op.levels; if ($levels -lt 1 -or $levels -gt 9) { throw "bad value: levels: $levels" } }
  $np = WdNewPara $doc $op.at $true
  $p = WdParaAt $doc $np.start
  $p.Range.Style = $doc.Styles.Item(-1); $p.Reset(); $p.Range.Font.Reset()
  $titleAt = 0
  if ($op.title) {
    $body = WdBody $p
    $body.Text = [string]$op.title
    $p = WdParaAt $doc $np.start
    $p.Range.Style = $doc.Styles.Item(-267)
    $titleAt = WdParaIndexAt $doc $np.start
    $p = WdParaAt $doc (WdSplitAfter $doc $p)
    $p.Range.Style = $doc.Styles.Item(-1); $p.Reset(); $p.Range.Font.Reset()
  }
  $start = [int]$p.Range.Start
  $r = $p.Range; $r.Collapse(1)
  # Range, UseHeadingStyles, UpperHeadingLevel, LowerHeadingLevel
  $toc = $doc.TablesOfContents.Add($r, $true, 1, $levels)
  $track = [bool]$doc.TrackRevisions
  if ($track) { $doc.TrackRevisions = $false }
  try { $toc.Update() } finally { if ($track) { $doc.TrackRevisions = $true } }
  # no heading paragraphs: Word fills in its own (English) 'No table of contents entries found.' - entries carry fields
  if ([int]$toc.Range.Fields.Count -le 1) { throw "no headings: $levels" }
  $res = @{ paragraph = (WdParaIndexAt $doc $start); levels = $levels; entries = [int]$toc.Range.Paragraphs.Count }
  if ($titleAt) { $res.title = $titleAt }
  return $res
}

# refreshed field results are not edits: keep them out of the tracked changes
function WdOpUpdateFields($doc, $word, $op) {
  $n = 0; $tocs = 0
  $track = [bool]$doc.TrackRevisions
  if ($track) { $doc.TrackRevisions = $false }
  try {
    foreach ($story in $doc.StoryRanges) {
      $r = $story
      while ($r) { $c = [int]$r.Fields.Count; if ($c -gt 0) { $n += $c; [void]$r.Fields.Update() }; $r = $r.NextStoryRange }
    }
    foreach ($toc in $doc.TablesOfContents) { $toc.Update(); $tocs++ }
    foreach ($tof in $doc.TablesOfFigures) { $tof.Update() }
  } finally { if ($track) { $doc.TrackRevisions = $true } }
  return @{ fields = $n; tocs = $tocs }
}

function WdOpAddFootnote($doc, $word, $op) {
  $hits = WdFind $doc $op.find ([bool]$op.all)
  for ($k = $hits.Count - 1; $k -ge 0; $k--) {
    $r = $hits[$k]; $r.Collapse(0)
    [void]$doc.Footnotes.Add($r, [Type]::Missing, [string]$op.text)
  }
  return @{ count = $hits.Count; total = [int]$doc.Footnotes.Count }
}

function WdOpInsertHyperlink($doc, $word, $op) {
  $url = [string]$op.url
  if ($url -notmatch '^(https?://|mailto:|ftp://)') { throw "bad value: url: $url" }
  $hits = WdFind $doc $op.find ([bool]$op.all)
  for ($k = $hits.Count - 1; $k -ge 0; $k--) { [void]$doc.Hyperlinks.Add($hits[$k], $url) }
  return @{ count = $hits.Count; total = [int]$doc.Hyperlinks.Count }
}

function WdEditOp($doc, $word, $op) {
  switch ([string]$op.op) {
    'insert_paragraph'   { return WdOpInsertParagraph $doc $word $op }
    'set_paragraph_text' { return WdOpSetParagraphText $doc $word $op }
    'delete_paragraph'   { return WdOpDeleteParagraph $doc $word $op }
    'apply_style'        { return WdOpApplyStyle $doc $word $op }
    'format_text'        { return WdOpFormatText $doc $word $op }
    'format_paragraph'   { return WdOpFormatParagraph $doc $word $op }
    'set_list'           { return WdOpSetList $doc $word $op }
    'insert_table'       { return WdOpInsertTable $doc $word $op }
    'table_set_cell'     { return WdOpTableSetCell $doc $word $op }
    'table_add_row'      { return WdOpTableAddRow $doc $word $op }
    'table_delete_row'   { return WdOpTableDeleteRow $doc $word $op }
    'insert_image'       { return WdOpInsertImage $doc $word $op }
    'insert_page_break'  { return WdOpInsertPageBreak $doc $word $op }
    'set_header'         { return WdOpHeaderFooter $doc $word $op $false }
    'set_footer'         { return WdOpHeaderFooter $doc $word $op $true }
    'page_setup'         { return WdOpPageSetup $doc $word $op }
    'add_comment'        { return WdOpAddComment $doc $word $op }
    'delete_comments'    { return WdOpDeleteComments $doc $word $op }
    'track_changes'      { $doc.TrackRevisions = [bool]$op.on; return @{ on = [bool]$doc.TrackRevisions } }
    'accept_revisions'   { $n = WdStoryRevisions $doc; $doc.AcceptAllRevisions(); return @{ count = $n; left = (WdStoryRevisions $doc) } }
    'reject_revisions'   { $n = WdStoryRevisions $doc; $doc.RejectAllRevisions(); return @{ count = $n; left = (WdStoryRevisions $doc) } }
    'insert_toc'         { return WdOpInsertToc $doc $word $op }
    'update_fields'      { return WdOpUpdateFields $doc $word $op }
    'add_footnote'       { return WdOpAddFootnote $doc $word $op }
    'insert_hyperlink'   { return WdOpInsertHyperlink $doc $word $op }
    default { throw "unknown op: $($op.op)" }
  }
}

# runs every operation in order; any failure throws before office.ps1 saves anything
function Invoke-WordEdit($doc, $word, $req) {
  # opened read-only (read-only file, or open in another program): Save() would wait on a hidden Save As dialog
  if ($req.save -and ([string]$req.save -eq [string]$req.path) -and [bool]$doc.ReadOnly) { throw 'read only' }
  $results = New-Object System.Collections.ArrayList
  $n = 0
  foreach ($op in @($req.operations)) {
    $n++
    $out = $null
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    try { $out = @(WdEditOp $doc $word $op) } catch {
      $msg = $_.Exception.Message
      if ($_.Exception.InnerException) { $msg = $_.Exception.InnerException.Message }
      throw ('op {0} ({1}): {2}' -f $n, [string]$op.op, $msg)
    }
    # the op result is the last object written (stray COM return values come before it)
    $r = $(if ($out.Count) { $out[$out.Count - 1] } else { @{} })
    if (-not ($r -is [hashtable])) { $r = @{} }
    $r.op = [string]$op.op
    $r.paragraphs = [int]$doc.Paragraphs.Count
    $r.ms = [int]$sw.ElapsedMilliseconds   # time inside Word (no lock wait): tests and diagnostics
    [void]$results.Add($r)
  }
  return @($results)
}
