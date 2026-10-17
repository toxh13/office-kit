# office.ps1 - drives installed Hancom/Word/PowerPoint/Excel through COM.
# Usage: powershell -File office.ps1 <request.json> <response.json>
# Request:  { "app": "hwp|word|ppt|excel|pdf|probe", "op": "...", "path": "...", "pidFile": "...", ... }
#           pdf: { "app": "pdf", "op": "pageImage", "path": "a.pdf", "page": 1, "dpi": 110, "output": "a.png" }
# Response: { "ok": true, ... } or { "ok": false, "error": "..." }
# Errors are short English codes ("open failed: ...", "unsupported format: ...") that com.mjs turns into Korean.
# The file stays ASCII: Windows PowerShell 5.1 reads BOM-less scripts in the ANSI code page.
param([string]$RequestPath, [string]$ResponsePath)
$ErrorActionPreference = 'Stop'

function Write-Response($obj) {
  $json = $obj | ConvertTo-Json -Depth 8 -Compress
  [System.IO.File]::WriteAllText($ResponsePath, $json, (New-Object System.Text.UTF8Encoding($false)))
}

function Release($o) { if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }

function Test-Prog($progId) { Test-Path "Registry::HKEY_CLASSES_ROOT\$progId\CLSID" }

# case-sensitive literal count (same rule as every replace below)
function Count-Occurrences([string]$text, [string]$find) {
  if (-not $find -or -not $text) { return 0 }
  return ([regex]::Matches($text, [regex]::Escape($find))).Count
}

# Office programs are shared per user session: one request at a time across all processes
# (the app preview and the AI's MCP server run in different processes).
$script:lock = $null
function Enter-OfficeLock {
  $script:lock = New-Object System.Threading.Mutex($false, 'Local\OfficeKitCom')
  try { [void]$script:lock.WaitOne() } catch [System.Threading.AbandonedMutexException] { }
}
function Exit-OfficeLock {
  if ($script:lock) { try { $script:lock.ReleaseMutex() } catch {}; $script:lock.Dispose(); $script:lock = $null }
}

# Office started through DCOM is not a child of this PowerShell: write its PID so Node can end it on timeout.
# Only a process that did not exist before is recorded (never the user's own running instance).
function New-OfficeApp([string]$progId, [string]$procName) {
  $before = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | ForEach-Object { $_.Id })
  $app = New-Object -ComObject $progId
  if ($req.pidFile) {
    $new = @(Get-Process -Name $procName -ErrorAction SilentlyContinue | Where-Object { $before -notcontains $_.Id } | ForEach-Object { $_.Id })
    if ($new.Count -eq 1) { $script:officePid = $new[0]; [System.IO.File]::WriteAllText($req.pidFile, [string]$new[0]) }
  }
  return $app
}

# after Quit, let the program we started really end before the next request (PowerPoint is single-instance)
$script:officePid = 0
function Wait-OfficeExit {
  if (-not $script:officePid) { return }
  try { [void](Get-Process -Id $script:officePid -ErrorAction Stop).WaitForExit(5000) } catch {}
}

# Documents opened by automation run macros unless told otherwise (msoAutomationSecurityLow is the default).
# Headless work: force-disable (3). Visible 'open' for the user: follow the Trust Center (2).
function Set-MacroGuard($app) {
  $prev = $null
  try { $prev = $app.AutomationSecurity; $app.AutomationSecurity = $(if ($req.visible) { 2 } else { 3 }) } catch {}
  return $prev
}
function Restore-MacroGuard($app, $prev) { if ($null -ne $prev) { try { $app.AutomationSecurity = $prev } catch {} } }

# allowed save targets per program (checked before the program starts)
$SaveFormats = @{
  hwp   = @{ hwp = 'HWP'; hwpx = 'HWPX'; hwt = 'HWP'; pdf = 'PDF'; docx = 'OOXML'; txt = 'TEXT'; html = 'HTML'; odt = 'ODT' }
  word  = @{ docx = 16; pdf = 17; doc = 0; rtf = 6; txt = 7; html = 10; odt = 23 }
  ppt   = @{ pptx = 24; pdf = 32; ppt = 1; odp = 35 }
  excel = @{ xlsx = 51; xlsm = 52; xls = 56; csv = 62; ods = 60; pdf = -1 }
}
function Get-SaveFormat([string]$app) {
  if (-not $req.save) { return $null }
  $fmt = $SaveFormats[$app][[string]$req.saveFormat]
  if ($null -eq $fmt) { throw "unsupported format: $($req.saveFormat)" }
  return $fmt
}

$NoPassword = 'officekit-no-password'
$Missing = [Type]::Missing

$req = [System.IO.File]::ReadAllText($RequestPath, [System.Text.Encoding]::UTF8) | ConvertFrom-Json

# ---------------------------------------------------------------- Hancom
function Invoke-Hwp($req) {
  $fmt = Get-SaveFormat 'hwp'
  $hwp = New-OfficeApp 'HWPFrame.HwpObject' 'Hwp'
  try {
    $secured = $false
    try { $secured = [bool]$hwp.RegisterModule('FilePathCheckDLL', $req.securityModule) } catch {}
    try { $hwp.XHwpWindows.Item(0).Visible = [bool]$req.visible } catch {}
    if ($req.op -ne 'create') {
      if (-not $hwp.Open($req.path, '', 'forceopen:true;suspendpassword:true')) { throw "open failed: $($req.path)" }
    }
    $res = @{ ok = $true; engine = 'hancom'; secured = $secured }
    switch ($req.op) {
      'text'   { $res.text = $hwp.GetTextFile('TEXT', '') }
      'fields' {
        $list = $hwp.GetFieldList(0, 0)
        $names = @(); if ($list) { $names = @($list -split [char]2 | Where-Object { $_ }) }
        $res.fields = @($names | ForEach-Object { @{ name = $_; value = $hwp.GetFieldText($_) } })
      }
      'fill' {
        $names = @(); $values = @()
        foreach ($p in $req.values.PSObject.Properties) { $names += $p.Name; $values += [string]$p.Value }
        $hwp.PutFieldText(($names -join [char]2), ($values -join [char]2))
        $res.filled = $names
      }
      'replace' {
        $res.count = 0; $counts = New-Object System.Collections.ArrayList
        foreach ($pair in $req.pairs) {
          $n = Count-Occurrences $hwp.GetTextFile('TEXT', '') $pair[0]
          if ($n -gt 0) {
            $set = $hwp.HParameterSet.HFindReplace
            [void]$hwp.HAction.GetDefault('AllReplace', $set.HSet)
            $set.FindString = $pair[0]
            $set.ReplaceString = $pair[1]
            $set.MatchCase = 1
            $set.IgnoreMessage = 1
            $set.Direction = 2          # whole document
            $set.FindType = 1
            [void]$hwp.HAction.Execute('AllReplace', $set.HSet)
          }
          [void]$counts.Add($n); $res.count += $n
        }
        $res.counts = @($counts)
      }
      'pageImage' {
        $res.pages = $hwp.PageCount
        [void]$hwp.CreatePageImage($req.output, [int]$req.page, [int]($req.dpi | ForEach-Object { if ($_) { $_ } else { 96 } }), 24, 'png')
        $res.output = $req.output
      }
    }
    if ($req.save) {
      if (-not $hwp.SaveAs($req.save, $fmt, '')) { throw "save failed: $($req.save)" }
      $res.saved = $req.save
    }
    if ($req.visible) { $res.leftOpen = $true } else { $hwp.Clear(1) }
    return $res
  } finally {
    if (-not $req.visible) { try { $hwp.Quit() } catch {} }
    Release $hwp
  }
}

# ---------------------------------------------------------------- Word
function Invoke-Word($req) {
  $fmt = Get-SaveFormat 'word'
  $word = New-OfficeApp 'Word.Application' 'WINWORD'
  $doc = $null; $guard = $null
  try {
    $word.DisplayAlerts = 0
    $guard = Set-MacroGuard $word
    if ($req.visible) { $word.Visible = $true }
    $readOnly = -not ($req.save -or $req.op -in @('replace', 'fill'))
    # a dummy password makes a protected document fail at once instead of waiting on a hidden prompt
    try { $doc = $word.Documents.Open($req.path, $false, $readOnly, $false, $NoPassword) }
    catch { throw "open failed: $($req.path) ($($_.Exception.Message))" }
    $res = @{ ok = $true; engine = 'word' }
    switch ($req.op) {
      'text'   { $res.text = $doc.Content.Text; $res.pages = $doc.ComputeStatistics(2) }
      'fields' {
        $res.fields = @()
        foreach ($cc in $doc.ContentControls) { $res.fields += @{ name = $(if ($cc.Title) { $cc.Title } else { $cc.Tag }); value = $cc.Range.Text; kind = 'contentControl' } }
        foreach ($bm in $doc.Bookmarks) { $res.fields += @{ name = $bm.Name; value = $bm.Range.Text; kind = 'bookmark' } }
      }
      'fill' {
        $res.filled = @()
        foreach ($p in $req.values.PSObject.Properties) {
          $hit = $false
          foreach ($cc in $doc.ContentControls) { if ($cc.Title -eq $p.Name -or $cc.Tag -eq $p.Name) { $cc.Range.Text = [string]$p.Value; $hit = $true } }
          if (-not $hit -and $doc.Bookmarks.Exists($p.Name)) {
            $r = $doc.Bookmarks.Item($p.Name).Range; $r.Text = [string]$p.Value; [void]$doc.Bookmarks.Add($p.Name, $r); $hit = $true
          }
          if ($hit) { $res.filled += $p.Name }
        }
      }
      'replace' {
        $res.count = 0; $counts = New-Object System.Collections.ArrayList
        foreach ($pair in $req.pairs) {
          $find = [string]$pair[0]; $repl = [string]$pair[1]
          if ($find.Length -gt 255 -or $repl.Length -gt 255) { throw 'find too long' }
          # count in every story (body, headers, footers, notes, text boxes) with the same case-sensitive rule
          $n = 0
          foreach ($story in $doc.StoryRanges) { $r = $story; while ($r) { $n += Count-Occurrences ([string]$r.Text) $find; $r = $r.NextStoryRange } }
          if ($n -gt 0) {
            # '^' starts a Word special code (^p, ^t, ^&): double it to keep it literal
            $wFind = $find.Replace('^', '^^'); $wRepl = $repl.Replace('^', '^^')
            foreach ($story in $doc.StoryRanges) {
              $r = $story
              while ($r) {
                $f = $r.Find; $f.ClearFormatting(); $f.Replacement.ClearFormatting()
                foreach ($opt in @('MatchByte', 'MatchFuzzy', 'IgnoreSpace', 'IgnorePunct', 'MatchPrefix', 'MatchSuffix')) {
                  try { $f.$opt = ($opt -eq 'MatchByte') } catch {}
                }
                # FindText, MatchCase, WholeWord, Wildcards, SoundsLike, AllForms, Forward, Wrap(stop), Format, ReplaceWith, Replace(all)
                [void]$f.Execute($wFind, $true, $false, $false, $false, $false, $true, 0, $false, $wRepl, 2)
                $r = $r.NextStoryRange
              }
            }
          }
          [void]$counts.Add($n); $res.count += $n
        }
        $res.counts = @($counts)
      }
    }
    if ($req.save) {
      if ($req.save -eq $req.path) { $doc.Save() }
      elseif ($fmt -eq 17) { $doc.ExportAsFixedFormat($req.save, 17) }
      else { $doc.SaveAs2($req.save, $fmt) }
      $res.saved = $req.save
    }
    if ($req.visible) { $res.leftOpen = $true; $doc = $null }
    return $res
  } finally {
    if ($doc) { try { $doc.Close($false) } catch {} }
    Restore-MacroGuard $word $guard
    if (-not $req.visible) { try { if ($word.Documents.Count -eq 0) { $word.Quit() } } catch {} }
    Release $doc; Release $word
  }
}

# ---------------------------------------------------------------- PowerPoint
function Invoke-Ppt($req) {
  $fmt = Get-SaveFormat 'ppt'
  $app = New-OfficeApp 'PowerPoint.Application' 'POWERPNT'
  $pres = $null; $guard = $null
  try {
    $app.DisplayAlerts = 1  # ppAlertsNone
    $guard = Set-MacroGuard $app
    $readOnly = $(if ($req.save -or $req.op -eq 'replace') { 0 } else { -1 })
    $withWindow = $(if ($req.visible) { -1 } else { 0 })
    try { $pres = $app.Presentations.Open($req.path, $readOnly, 0, $withWindow) }
    catch { throw "open failed: $($req.path) ($($_.Exception.Message))" }
    $res = @{ ok = $true; engine = 'powerpoint'; slides = $pres.Slides.Count }
    $frames = New-Object System.Collections.ArrayList
    foreach ($slide in $pres.Slides) {
      foreach ($shape in $slide.Shapes) {
        if ($shape.HasTextFrame) { [void]$frames.Add(@{ slide = $slide.SlideIndex; tr = $shape.TextFrame.TextRange }) }
        if ($shape.HasTable) {
          foreach ($row in $shape.Table.Rows) { foreach ($cell in $row.Cells) { [void]$frames.Add(@{ slide = $slide.SlideIndex; tr = $cell.Shape.TextFrame.TextRange }) } }
        }
      }
    }
    switch ($req.op) {
      'text' {
        $bySlide = @{}
        foreach ($f in $frames) { $bySlide[$f.slide] += ($f.tr.Text -replace "`r", "`n") + "`n" }
        $res.text = (1..$pres.Slides.Count | ForEach-Object { "--- slide $_ ---`n" + $bySlide[$_] }) -join "`n"
      }
      'replace' {
        $count = 0; $counts = New-Object System.Collections.ArrayList
        foreach ($pair in $req.pairs) {
          $n = 0
          foreach ($f in $frames) {
            $tr = $f.tr; $after = 0
            while ($true) {
              # FindWhat, ReplaceWhat, After, MatchCase(msoTrue), WholeWords(msoFalse)
              $hit = $tr.Replace($pair[0], $pair[1], $after, -1, 0)
              if (-not $hit) { break }
              $n++; $after = $hit.Start + $hit.Length - 1
            }
          }
          [void]$counts.Add($n); $count += $n
        }
        $res.count = $count; $res.counts = @($counts)
      }
      'slideImage' {
        $pres.Slides.Item([int]$req.page).Export($req.output, 'PNG')
        $res.output = $req.output
      }
    }
    if ($req.save) {
      if ($req.save -eq $req.path) { $pres.Save() } else { $pres.SaveAs($req.save, $fmt) }
      $res.saved = $req.save
    }
    if ($req.visible) { $res.leftOpen = $true; $pres = $null }
    return $res
  } finally {
    if ($pres) { try { $pres.Close() } catch {} }
    Restore-MacroGuard $app $guard
    if (-not $req.visible) { try { if ($app.Presentations.Count -eq 0) { $app.Quit() } } catch {} }
    Release $pres; Release $app
  }
}

# ---------------------------------------------------------------- Excel
$XlErrors = @{ -2146826281 = '#DIV/0!'; -2146826246 = '#N/A'; -2146826259 = '#NAME?'; -2146826288 = '#NULL!';
  -2146826252 = '#NUM!'; -2146826265 = '#REF!'; -2146826273 = '#VALUE!' }

function Format-Cell($v) {
  if ($null -eq $v) { return '' }
  if ($v -is [string]) { return $v }
  if ($v -is [double]) { return $v.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
  if ($v -is [bool]) { return $(if ($v) { 'TRUE' } else { 'FALSE' }) }
  if ($v -is [datetime]) { if ($v.TimeOfDay.Ticks -eq 0) { return $v.ToString('yyyy-MM-dd') } else { return $v.ToString('yyyy-MM-dd HH:mm:ss') } }
  if ($v -is [int]) { $e = $XlErrors[$v]; if ($e) { return $e } }
  return [string]$v
}

# Range.Replace treats * ? ~ as wildcards: escape them with ~
function Escape-ExcelFind([string]$s) { return $s.Replace('~', '~~').Replace('*', '~*').Replace('?', '~?') }

function Invoke-Excel($req) {
  $fmt = Get-SaveFormat 'excel'
  $app = New-OfficeApp 'Excel.Application' 'EXCEL'
  $wb = $null; $guard = $null
  try {
    $app.DisplayAlerts = $false
    $guard = Set-MacroGuard $app
    if ($req.visible) { $app.Visible = $true } else { try { $app.EnableEvents = $false } catch {} }
    $readOnly = -not ($req.save -or $req.op -in @('replace', 'setCells'))
    # Filename, UpdateLinks(never), ReadOnly, Format, Password(dummy), WriteResPassword, IgnoreReadOnlyRecommended
    try { $wb = $app.Workbooks.Open($req.path, 0, $readOnly, $Missing, $NoPassword, $Missing, $true) }
    catch { throw "open failed: $($req.path) ($($_.Exception.Message))" }
    $res = @{ ok = $true; engine = 'excel' }
    switch ($req.op) {
      'text' {
        # one read per sheet (cell-by-cell COM calls are far too slow); rows/columns from A1 like the file engine
        $maxRows = 10000; $maxCols = 200; $maxCells = 200000
        $sheets = New-Object System.Collections.ArrayList
        foreach ($ws in $wb.Worksheets) {
          $ur = $ws.UsedRange
          $last = $ur.Cells.Item($ur.Rows.Count, $ur.Columns.Count)
          $totalRows = [int]$last.Row; $totalCols = [int]$last.Column
          $useCols = [Math]::Min($totalCols, $maxCols)
          $useRows = [Math]::Min($totalRows, [Math]::Min($maxRows, [Math]::Max(1, [Math]::Floor($maxCells / $useCols))))
          $rg = $ws.Range($ws.Cells.Item(1, 1), $ws.Cells.Item($useRows, $useCols))
          $vals = $null
          try { $vals = $rg.Value(10) } catch { $vals = $rg.Value2 }
          # ::new, not New-Object: a PSObject-wrapped array serializes as {value, Count}
          $rows = [object[]]::new($useRows)
          if ($vals -is [System.Array] -and $vals.Rank -eq 2) {
            $r0 = $vals.GetLowerBound(0); $c0 = $vals.GetLowerBound(1)
            for ($r = 0; $r -lt $useRows; $r++) {
              $row = [string[]]::new($useCols)
              for ($c = 0; $c -lt $useCols; $c++) { $row[$c] = Format-Cell $vals[($r + $r0), ($c + $c0)] }
              $rows[$r] = $row
            }
          } else {
            $rows[0] = [string[]]@(Format-Cell $vals)
          }
          [void]$sheets.Add(@{ name = $ws.Name; address = $ur.Address($false, $false); rows = $rows;
            totalRows = $totalRows; totalCols = $totalCols; truncated = ($useRows -lt $totalRows -or $useCols -lt $totalCols) })
        }
        $res.sheets = @($sheets)
      }
      'replace' {
        $count = 0; $counts = New-Object System.Collections.ArrayList
        foreach ($pair in $req.pairs) {
          $find = [string]$pair[0]; $n = 0
          foreach ($ws in $wb.Worksheets) {
            $ur = $ws.UsedRange
            # Replace works on formulas/constants: count over the same text, read in one call
            $m = 0
            foreach ($v in @($ur.Formula)) { if ($v -is [string]) { $m += Count-Occurrences $v $find } }
            # What, Replacement, LookAt(xlPart), SearchOrder(xlByRows), MatchCase
            if ($m -gt 0) { [void]$ur.Replace((Escape-ExcelFind $find), [string]$pair[1], 2, 1, $true) }
            $n += $m
          }
          [void]$counts.Add($n); $count += $n
        }
        $res.count = $count; $res.counts = @($counts)
      }
      'setCells' {
        $ws = $null
        if ($req.sheet) { try { $ws = $wb.Worksheets.Item([string]$req.sheet) } catch { throw "sheet not found: $($req.sheet)" } }
        else { $ws = $wb.Worksheets.Item(1) }
        $n = 0
        foreach ($p in $req.cells.PSObject.Properties) {
          $v = [string]$p.Value
          if ($v.StartsWith('=')) { $ws.Range($p.Name).Formula = $v } else { $ws.Range($p.Name).Value2 = $p.Value }
          $n++
        }
        $res.updated = $n
      }
    }
    if ($req.save) {
      if ($req.save -eq $req.path) { $wb.Save() }
      elseif ($fmt -eq -1) { $wb.ExportAsFixedFormat(0, $req.save) }
      else { $wb.SaveAs($req.save, $fmt) }
      $res.saved = $req.save
    }
    if ($req.visible) { $res.leftOpen = $true; $wb = $null }
    return $res
  } finally {
    if ($wb) { try { $wb.Close($false) } catch {} }
    Restore-MacroGuard $app $guard
    if (-not $req.visible) { try { if ($app.Workbooks.Count -eq 0) { $app.Quit() } } catch {} }
    Release $wb; Release $app
  }
}

# ---------------------------------------------------------------- PDF (Windows.Data.Pdf, no Office needed)
function Await-WinRT($op, [Type]$resultType) {
  $m = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncOperation`1' } | Select-Object -First 1
  $task = $m.MakeGenericMethod($resultType).Invoke($null, @($op))
  [void]$task.Wait(-1)
  return $task.Result
}

function Await-WinRTAction($action) {
  $m = [System.WindowsRuntimeSystemExtensions].GetMethods() | Where-Object {
    $_.Name -eq 'AsTask' -and -not $_.IsGenericMethod -and $_.GetParameters().Count -eq 1 -and
    $_.GetParameters()[0].ParameterType.Name -eq 'IAsyncAction' } | Select-Object -First 1
  [void]$m.Invoke($null, @($action)).Wait(-1)
}

function Invoke-Pdf($req) {
  if ($req.op -ne 'pageImage') { throw "unknown pdf op: $($req.op)" }
  Add-Type -AssemblyName System.Runtime.WindowsRuntime
  $null = [Windows.Data.Pdf.PdfDocument, Windows.Data.Pdf, ContentType = WindowsRuntime]
  $null = [Windows.Data.Pdf.PdfPageRenderOptions, Windows.Data.Pdf, ContentType = WindowsRuntime]
  $null = [Windows.Storage.Streams.InMemoryRandomAccessStream, Windows.Storage.Streams, ContentType = WindowsRuntime]
  $file = [System.IO.File]::OpenRead($req.path)
  $doc = $null; $pdfPage = $null; $mem = $null
  try {
    $ras = [System.IO.WindowsRuntimeStreamExtensions]::AsRandomAccessStream($file)
    try { $doc = Await-WinRT ([Windows.Data.Pdf.PdfDocument]::LoadFromStreamAsync($ras)) ([Windows.Data.Pdf.PdfDocument]) }
    catch { throw "open failed: $($req.path)" }
    $res = @{ ok = $true; engine = 'pdf'; pages = [int]$doc.PageCount; page = [int]$req.page }
    # out of range: report page count only (caller builds the message)
    if ([int]$req.page -lt 1 -or [int]$req.page -gt $doc.PageCount) { $res.output = $null; return $res }
    $pdfPage = $doc.GetPage([uint32]([int]$req.page - 1))
    $dpi = $(if ($req.dpi) { [double]$req.dpi } else { 96.0 })
    $wantW = [Math]::Max(1, [Math]::Round($pdfPage.Size.Width * $dpi / 96.0))
    $wantH = [Math]::Max(1, [Math]::Round($pdfPage.Size.Height * $dpi / 96.0))
    # DestinationWidth is in DIPs, so display scaling (e.g. 125%) enlarges the bitmap: render, measure, re-render once
    $scale = 1.0; $bytes = $null
    for ($try = 0; $try -lt 2; $try++) {
      $opts = New-Object Windows.Data.Pdf.PdfPageRenderOptions
      $opts.DestinationWidth = [uint32][Math]::Max(1, [Math]::Round($wantW / $scale))
      $opts.DestinationHeight = [uint32][Math]::Max(1, [Math]::Round($wantH / $scale))
      if ($mem) { $mem.Dispose() }
      $mem = New-Object Windows.Storage.Streams.InMemoryRandomAccessStream
      Await-WinRTAction ($pdfPage.RenderToStreamAsync($mem, $opts))
      $src = [System.IO.WindowsRuntimeStreamExtensions]::AsStreamForRead($mem.GetInputStreamAt(0))
      $ms = New-Object System.IO.MemoryStream
      $src.CopyTo($ms); $bytes = $ms.ToArray(); $ms.Dispose()
      # PNG IHDR width at bytes 16..19 (big-endian)
      $gotW = ([int]$bytes[16] -shl 24) -bor ([int]$bytes[17] -shl 16) -bor ([int]$bytes[18] -shl 8) -bor [int]$bytes[19]
      $gotH = ([int]$bytes[20] -shl 24) -bor ([int]$bytes[21] -shl 16) -bor ([int]$bytes[22] -shl 8) -bor [int]$bytes[23]
      if ([Math]::Abs($gotW - $wantW) -le 2 -or $gotW -le 0) { break }
      $scale = $scale * $gotW / [double]$opts.DestinationWidth
    }
    [System.IO.File]::WriteAllBytes($req.output, $bytes)
    $res.output = $req.output
    $res.width = $gotW; $res.height = $gotH
    return $res
  } finally {
    if ($mem) { $mem.Dispose() }
    if ($pdfPage) { $pdfPage.Dispose() }
    $file.Dispose()
  }
}

try {
  switch ($req.app) {
    'probe' {
      $mods = $null
      try { $mods = Get-ItemProperty 'HKCU:\Software\HNC\HwpAutomation\Modules' -ErrorAction Stop } catch {}
      Write-Response @{ ok = $true; hancom = (Test-Prog 'HWPFrame.HwpObject'); word = (Test-Prog 'Word.Application');
        powerpoint = (Test-Prog 'PowerPoint.Application'); excel = (Test-Prog 'Excel.Application');
        hancomSecurityModule = [bool]($mods -and $mods.PSObject.Properties[$req.securityModule]) }
    }
    'pdf'   { Write-Response (Invoke-Pdf $req) }
    { $_ -in @('hwp', 'word', 'ppt', 'excel') } {
      Enter-OfficeLock
      $out = $null
      switch ($req.app) {
        'hwp'   { $out = Invoke-Hwp $req }
        'word'  { $out = Invoke-Word $req }
        'ppt'   { $out = Invoke-Ppt $req }
        'excel' { $out = Invoke-Excel $req }
      }
      Write-Response $out
    }
    default { throw "unknown app: $($req.app)" }
  }
} catch {
  Write-Response @{ ok = $false; error = $_.Exception.Message }
  exit 1
} finally {
  [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if (-not $req.visible) { Wait-OfficeExit }
  Exit-OfficeLock
}
