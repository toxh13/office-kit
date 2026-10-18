# strict-open.ps1 - opens a docx/xlsx/pptx in Word/Excel/PowerPoint the strict way and reports whether it opened cleanly.
# Usage: powershell -NoProfile -ExecutionPolicy Bypass -File strict-open.ps1 -Path <file> [-Out <result.json>] [-TimeoutSec 60] [-LockWaitSec 300]
#
#   - ALL alerts on (Word wdAlertsAll, Excel DisplayAlerts, PowerPoint ppAlertsAll), repair off (Word OpenAndRepair=false,
#     Excel CorruptLoad=xlNormalLoad), read-only, macros force-disabled, a dummy password so a protected file fails instead of prompting.
#   - A file Office wants to repair makes it show a modal prompt, so the open call never returns. A watchdog (separate
#     process) then kills exactly the Office process this script started (matched by PID and start time) -> "hung".
#   - Never touches the user's own Office: PowerPoint is single-instance, so if one is already running we refuse (exit 3)
#     instead of attaching to it; any instance that existed before this script is never quit or killed.
#   - Takes Office Kit's shared lock (Local\OfficeKitCom, same as scripts/com/office.ps1) so it never races Office Kit's own COM work.
#
# Result JSON (UTF-8): { file, app, ok, opened, hung, busy, pid, openMs, error, watchdog, prompt, forcedQuit, ...details }
#   prompt (only when hung): [{title, class, texts[]}] - the blocking windows of the killed process, read via UI Automation
#   word:  paragraphs, paras[{text,bold,font}], tables[{rows,columns,uniform,rowsItem,rowsItemError,cells[]}]
#   excel: sheets[{name,used,rows[]}]
#   ppt:   slides[{index,shapes[{name,type,text,bold}]}]
# Exit code: 0 opened cleanly, 1 hung (prompt) or failed, 2 bad arguments, 3 busy (lock timeout / user's PowerPoint running).
# The file stays ASCII: Windows PowerShell 5.1 reads BOM-less scripts in the ANSI code page.
param([string]$Path, [string]$Out, [int]$TimeoutSec = 60, [int]$LockWaitSec = 300)
$ErrorActionPreference = 'Stop'

$res = [ordered]@{ file = $Path; app = $null; ok = $false; opened = $false; hung = $false; busy = $false }
$utf8 = New-Object System.Text.UTF8Encoding($false)
function Write-Result([int]$code) {
  $json = $res | ConvertTo-Json -Depth 8
  if ($Out) { [System.IO.File]::WriteAllText($Out, $json, $utf8) }
  else { [Console]::OutputEncoding = $utf8; [Console]::Out.WriteLine($json) }
  exit $code
}

$apps = @{
  '.docx' = @{ name = 'word';  progId = 'Word.Application';       proc = 'WINWORD' }
  '.xlsx' = @{ name = 'excel'; progId = 'Excel.Application';      proc = 'EXCEL' }
  '.pptx' = @{ name = 'ppt';   progId = 'PowerPoint.Application'; proc = 'POWERPNT' }
}
if (-not $Path -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { $res.error = "file not found: $Path"; Write-Result 2 }
$Path = (Resolve-Path -LiteralPath $Path).ProviderPath
$res.file = $Path
$spec = $apps[[System.IO.Path]::GetExtension($Path).ToLower()]
if (-not $spec) { $res.error = "unsupported file type: $Path"; Write-Result 2 }
$res.app = $spec.name

Add-Type @"
using System; using System.Runtime.InteropServices;
public static class StrictOpenW32 {
  [DllImport("user32.dll", SetLastError=true)] public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint pid);
  [DllImport("user32.dll", SetLastError=true, CharSet=CharSet.Unicode)] public static extern IntPtr FindWindow(string cls, string title);
}
"@
function PidOf([IntPtr]$h) { if ($h -eq [IntPtr]::Zero) { return 0 }; [uint32]$p = 0; [void][StrictOpenW32]::GetWindowThreadProcessId($h, [ref]$p); return [int]$p }
function Pids([string]$name) { @(Get-Process -Name $name -ErrorAction SilentlyContinue | ForEach-Object { $_.Id }) }
function Release($o) { if ($o) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch {} } }
$Missing = [Type]::Missing

# Each opener runs in its own function scope so every COM reference it made is gone (and collectable) when it returns:
# a reference left in a script variable keeps Office alive after Quit.

function Open-Word($app) {
  $app.DisplayAlerts = -1                         # wdAlertsAll
  $app.Visible = $false
  $t0 = Get-Date
  # FileName, ConfirmConversions, ReadOnly, AddToRecentFiles, PasswordDocument, PasswordTemplate, Revert,
  # WritePasswordDocument, WritePasswordTemplate, Format, Encoding, Visible, OpenAndRepair, DocumentDirection, NoEncodingDialog
  $docs = $app.Documents
  [object[]]$a = @([string]$Path, $false, $true, $false, 'officekit-no-password', $Missing, $false, $Missing, $Missing, 0, $Missing, $false, $false, $Missing, $true)
  $doc = $docs.GetType().InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null, $docs, $a)
  $res.openMs = [int]((Get-Date) - $t0).TotalMilliseconds
  $res.opened = $true
  $res.paragraphs = $doc.Paragraphs.Count
  $paras = @()
  $n = [Math]::Min($doc.Paragraphs.Count, 80)
  for ($i = 1; $i -le $n; $i++) {
    $r = $doc.Paragraphs.Item($i).Range
    $txt = ($r.Text -replace "[\r\a]", '')
    if ($txt.Length -gt 0) { $paras += [ordered]@{ text = $txt; bold = $r.Font.Bold; font = $r.Font.Name } }
  }
  $res.paras = $paras
  $tables = @()
  $tn = [Math]::Min($doc.Tables.Count, 10)
  for ($i = 1; $i -le $tn; $i++) {
    $t = $doc.Tables.Item($i)
    $info = [ordered]@{ rows = $t.Rows.Count; columns = $t.Columns.Count; uniform = $t.Uniform; rowsItem = $true }
    # Rows.Item(n) fails on a table Word thinks has vertically merged cells (Office Kit's word_edit table ops use it)
    try { for ($r = 1; $r -le $t.Rows.Count; $r++) { [void]$t.Rows.Item($r).Cells.Count } } catch { $info.rowsItem = $false; $info.rowsItemError = $_.Exception.Message }
    $cells = @()
    foreach ($c in $t.Range.Cells) { if ($cells.Count -lt 60) { $cells += ('r' + $c.RowIndex + 'c' + $c.ColumnIndex + '=' + ($c.Range.Text -replace "[\r\a]", '')) } }
    $info.cells = $cells
    $tables += $info
  }
  $res.tables = $tables
  $doc.Close($false)
}

function Open-Excel($app) {
  $app.DisplayAlerts = $true
  $app.Visible = $false
  $t0 = Get-Date
  # FileName, UpdateLinks, ReadOnly, Format, Password, WriteResPassword, IgnoreReadOnlyRecommended, Origin, Delimiter,
  # Editable, Notify, Converter, AddToMru, Local, CorruptLoad (0 = xlNormalLoad: no repair)
  $wbs = $app.Workbooks
  [object[]]$a = @([string]$Path, 0, $true, $Missing, 'officekit-no-password', $Missing, $true, $Missing, $Missing, $false, $false, $Missing, $false, $Missing, 0)
  $wb = $wbs.GetType().InvokeMember('Open', [Reflection.BindingFlags]::InvokeMethod, $null, $wbs, $a)
  $res.openMs = [int]((Get-Date) - $t0).TotalMilliseconds
  $res.opened = $true
  $sheets = @()
  foreach ($ws in $wb.Worksheets) {
    $used = $ws.UsedRange
    $rows = @()
    for ($r = 1; $r -le [Math]::Min($used.Rows.Count + $used.Row - 1, 20); $r++) {
      $vals = @()
      for ($c = 1; $c -le [Math]::Min($used.Columns.Count + $used.Column - 1, 8); $c++) {
        $cell = $ws.Cells.Item($r, $c)
        $v = [string]$cell.Text
        if ($cell.HasFormula) { $v += ' {' + $cell.Formula + '}' }
        if ($cell.Font.Bold -eq $true -and $v) { $v = '**' + $v + '**' }
        $vals += $v
      }
      $rows += ($vals -join ' | ')
    }
    $sheets += [ordered]@{ name = $ws.Name; used = $used.Address(); rows = $rows }
  }
  $res.sheets = $sheets
  $wb.Close($false)
}

function Open-Ppt($app) {
  $app.DisplayAlerts = 2                          # ppAlertsAll
  $t0 = Get-Date
  # FileName, ReadOnly, Untitled, WithWindow
  $pres = $app.Presentations.Open([string]$Path, -1, 0, 0)
  $res.openMs = [int]((Get-Date) - $t0).TotalMilliseconds
  $res.opened = $true
  $slides = @()
  foreach ($s in $pres.Slides) {
    $shapes = @()
    foreach ($sh in $s.Shapes) {
      $txt = ''; $b = $null
      if ($sh.HasTextFrame -and $sh.TextFrame.HasText) { $txt = $sh.TextFrame.TextRange.Text; $b = $sh.TextFrame.TextRange.Font.Bold }
      if ($sh.HasTable) { $txt = '[table ' + $sh.Table.Rows.Count + 'x' + $sh.Table.Columns.Count + ']' }
      $shapes += [ordered]@{ name = $sh.Name; type = $sh.Type; text = $txt; bold = $b }
    }
    $slides += [ordered]@{ index = $s.SlideIndex; shapes = $shapes }
  }
  $res.slides = $slides
  $pres.Close()
}

# ---------------------------------------------------------------- shared Office lock
$lock = New-Object System.Threading.Mutex($false, 'Local\OfficeKitCom')
$held = $false
try { $held = $lock.WaitOne([Math]::Max(1, $LockWaitSec) * 1000) } catch [System.Threading.AbandonedMutexException] { $held = $true }
if (-not $held) { $lock.Dispose(); $res.busy = $true; $res.error = "busy: another Office job held Local\OfficeKitCom for more than $LockWaitSec s"; Write-Result 3 }

$app = $null; $appPid = 0; $startTicks = 0; $job = $null; $ours = $false
try {
  $before = Pids $spec.proc
  if ($spec.name -eq 'ppt' -and $before.Count) {
    # an instance that just quit (Office Kit's own job) may still be ending
    $deadline = (Get-Date).AddSeconds(10)
    while ((Pids $spec.proc).Count -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }
    $before = Pids $spec.proc
    if ($before.Count) { $res.busy = $true; throw "busy: PowerPoint is already running (pid $($before -join ',')); strict open needs its own instance" }
  }

  $app = New-Object -ComObject $spec.progId
  if ($spec.name -eq 'word') {
    $tag = 'officekit-strict-' + [guid]::NewGuid().ToString('N')
    $app.Caption = $tag
    $appPid = PidOf ([StrictOpenW32]::FindWindow('OpusApp', $tag))
  } elseif ($spec.name -eq 'excel') {
    $appPid = PidOf ([IntPtr][long]$app.Hwnd)
  } else {
    $appPid = PidOf ([IntPtr][long]$app.HWND)
  }
  $new = @(Pids $spec.proc | Where-Object { $before -notcontains $_ })
  if (-not $appPid -and $new.Count -eq 1) { $appPid = $new[0] }
  $res.pid = $appPid
  if ($appPid -and $before -contains $appPid) {
    # attached to an instance that was already running: not ours -> drop the reference, never quit or kill it
    $res.busy = $true; Release $app; $app = $null
    throw "busy: COM attached to an already running $($spec.proc) (pid $appPid)"
  }
  if (-not $appPid) { $ours = ($new.Count -gt 0); throw "cannot tell which $($spec.proc) process COM started" }
  $ours = $true
  $startTicks = (Get-Process -Id $appPid).StartTime.Ticks

  # watchdog: kills only our process (same PID and start time) if the work does not return in time.
  # Before killing it records what was blocking: that process's top-level windows and their text (UI Automation, read only).
  $job = Start-Job -ArgumentList $appPid, $startTicks, $TimeoutSec -ScriptBlock {
    param($p, $ticks, $s)
    Start-Sleep -Seconds $s
    $out = [ordered]@{ result = 'gone' }
    $proc = Get-Process -Id $p -ErrorAction SilentlyContinue
    if ($proc -and $proc.StartTime.Ticks -eq $ticks) {
      try {
        Add-Type -AssemblyName UIAutomationClient, UIAutomationTypes
        $A = [System.Windows.Automation.AutomationElement]
        $cond = New-Object System.Windows.Automation.PropertyCondition($A::ProcessIdProperty, [int]$p)
        $wins = @()
        foreach ($w in $A::RootElement.FindAll([System.Windows.Automation.TreeScope]::Children, $cond)) {
          $texts = @()
          foreach ($e in $w.FindAll([System.Windows.Automation.TreeScope]::Descendants, [System.Windows.Automation.Condition]::TrueCondition)) {
            $n = $e.Current.Name
            if ($n -and $texts.Count -lt 30 -and $texts -notcontains $n) { $texts += $n }
          }
          $wins += [ordered]@{ title = $w.Current.Name; class = $w.Current.ClassName; texts = $texts }
          if ($wins.Count -ge 5) { break }
        }
        $out.windows = $wins
      } catch { $out.windowsError = $_.Exception.Message }
      try { Stop-Process -Id $p -Force -ErrorAction Stop; $out.result = 'killed' } catch { $out.result = 'kill failed' }
    }
    $out | ConvertTo-Json -Depth 5 -Compress
  }

  try { $app.AutomationSecurity = 3 } catch {}   # msoAutomationSecurityForceDisable
  if ($spec.name -eq 'word') { Open-Word $app } elseif ($spec.name -eq 'excel') { Open-Excel $app } else { Open-Ppt $app }
  $res.ok = $true
} catch {
  $res.error = $_.Exception.Message
} finally {
  if ($job) {
    if ((Get-Job -Id $job.Id).State -eq 'Running') { Stop-Job -Id $job.Id }
    else {
      $wd = [string](Receive-Job -Id $job.Id | Select-Object -Last 1)
      try {
        $w = $wd | ConvertFrom-Json
        $res.watchdog = $w.result
        if ($w.windows) { $res.prompt = @($w.windows) }
        if ($w.windowsError) { $res.promptError = $w.windowsError }
      } catch { $res.watchdog = $wd }
      $res.hung = ($res.watchdog -eq 'killed')
    }
    Remove-Job -Id $job.Id -Force
  }
  if ($res.hung) {
    $res.ok = $false
    $where = $(if ($res.opened) { 'after Open (reading the document)' } else { 'from Open (a prompt, e.g. repair)' })
    $res.error = "hung: $($spec.proc) did not return $where within $TimeoutSec s -> killed pid $appPid"
  }
  if ($ours -and $app -and -not $res.hung) { try { $app.Quit() } catch {} }
  Release $app
  $app = $null
  [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect(); [GC]::WaitForPendingFinalizers()
  if ($ours -and $appPid) {
    $proc = Get-Process -Id $appPid -ErrorAction SilentlyContinue
    if ($proc -and $proc.StartTime.Ticks -eq $startTicks) {
      if (-not $proc.WaitForExit(15000)) { try { Stop-Process -Id $appPid -Force; $res.forcedQuit = $true } catch {} }
    }
  }
  try { $lock.ReleaseMutex() } catch {}
  $lock.Dispose()
}
if ($res.ok) { Write-Result 0 } elseif ($res.busy) { Write-Result 3 } else { Write-Result 1 }
