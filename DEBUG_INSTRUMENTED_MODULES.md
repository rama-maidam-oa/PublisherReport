# Debug Instrumentation — OISR-21439 hang at 84% ("Applying Conditional Formats..")

## Logging behavior — OISR-22564 follow-up

The production template retains diagnostic checkpoints in memory and writes
`PublisherReport_Debug_<workbook>.log` only after an unexpected error is reported.
The `DragTo*` properties are intentionally not applicable to pivot data fields; skipping those
properties is recorded as a buffered warning, not an error, so a successful run must not create
a log file for that expected Excel limitation.

## Round 1 findings (from `PublisherReport_Debug.log`, run 2026-08-15 20:16)

The first instrumented run captured the stall directly. Relevant lines:

```
20:16:33 | +0.297s | CF.PivotCF | field=[SRG Unit Price ] area 1 BEFORE FormatConditions.Add
20:16:33 | +0.297s | CF.PivotCF | field=[SRG Unit Price ] area 1 AFTER FormatConditions.Add (Err=0 )
20:16:47 | +14.383s | CF.PivotCF | field=[SRG Unit Price ] area 1 font/fill/style applied (Err=-2147417848)
```

`-2147417848` = `0x8001010A` = **`RPC_E_SERVERCALL_RETRYLATER`** — a genuine COM-level "the
application is busy" error, not a VBA logic mistake. Key takeaways:

- `.FormatConditions.Add` itself is instant (`Err=0`, no time gap).
- The **~14 second stall happens somewhere inside the block that sets
  `.Font.Color` / `.Interior.Color` / `.Font.Bold` / `.Font.Italic` / `.Font.Underline`** on the
  newly-created condition — that's the only code between the two log lines.
- Both logged rules in this run had `underline=N` — so the `.Font.Underline` enum fix from the
  previous pass **is not what's causing this stall**; the delay reproduces on ordinary
  fill-color + bold + italic alone. The Underline fix should stay (it's still a real latent bug
  for reports that *do* use underline), but it is not the explanation for the hang.
- This points at something external holding Excel's message loop busy while that property
  write happens — most plausibly another COM add-in reacting to the change (the ribbon in your
  screenshots shows **Orbit GLSense**, **Orbit XLEdge**, **Acrobat**, **Copilot** all loaded
  alongside this workbook), or Excel itself internally re-rendering the Conditional Formatting
  engine. In a full production report with many more fields × rules, this same per-call stall
  compounding is almost certainly what turns into the multi-hour hang.

### Two things to do in parallel

**A. Quick, no-code-change test — isolate whether an add-in is the cause.** Before re-running
with new logging, try this once: `File → Options → Add-ins → Manage: COM Add-ins → Go...`,
uncheck **all** non-essential add-ins (Orbit GLSense, Orbit XLEdge, Copilot, Acrobat — leave only
what's strictly required), restart Excel, and reproduce the hang once more with the *current*
instrumented modules already installed. If the 14-second-per-rule stall disappears with add-ins
off, we've found it — the fix becomes "don't run this macro with those add-ins loaded" (or
report it to whichever add-in is responsible) rather than another CF.bas change. If it still
stalls with everything disabled, the cause is internal to Excel/the pivot itself, not an add-in.

**B. Finer-grained logging — isolate the exact property.** The updated `CF.bas` below adds a
log line **before and after each individual property assignment** (`Font.Color`,
`Interior.Color`, `Font.Bold`, `Font.Italic`, `Font.Underline`) instead of one log around all
five, with `Err.Clear` before each so the captured error actually belongs to that line. Re-paste
this into the `CF` module (no changes needed to `DebugLog.bas` or `MainModule.bas` — they're
already correct from the last pass) and reproduce the hang again. Send me the new log tail —
this time we'll see exactly which single property line eats the 14+ seconds.

## Round 2 findings (from the per-property log, run 2026-08-15 20:24) — root cause confirmed

```
+0.313s | BEFORE Font.Bold
+0.313s | AFTER Font.Bold (Err=0 )
+0.313s | BEFORE Font.Italic
+0.313s | AFTER Font.Italic (Err=-2147417848 Method 'Italic' of object 'Font' failed)
...
+0.313s | BEFORE Interior.Color        <- log stops here; Excel closed itself this run
```

`.Font.Bold = True` succeeds. The **very next** call — `.Font.Italic = True` on that same
`FormatCondition.Font` object — fails immediately with the same `RPC_E_SERVERCALL_RETRYLATER`.
Then, on the *second* conditional-format rule, the same pattern recurs and this time Excel
doesn't recover from the busy state at all — it closes itself while sitting at `Interior.Color`.

This is the actual mechanism behind OISR-21439: **writing two-plus font/fill properties
back-to-back on the same COM object races Excel's own internal handling of the first write.**
It matches the ticket's own description precisely ("triggered when two or more font properties
applied simultaneously... single formats work correctly") — and explains why Round 1 saw a
14-second delay-then-recover on one run and this run saw an instant failure escalating into a
full close: Excel's *own*, undocumented, opaque retry behavior for this COM error is what's
inconsistent, not our code's logic.

### The fix

Instead of relying on Excel's own opaque retry/give-up behavior, each property write gets an
explicit, bounded retry: on `RPC_E_SERVERCALL_RETRYLATER` specifically, call `DoEvents` (let
Excel's message loop catch up) plus a short `Sleep`, then retry, up to a small fixed number of
attempts. This bounds the worst case to a couple of seconds per property (instead of an
unbounded stall or a crash) and, in the normal case where the race doesn't occur, adds no
overhead at all (first attempt succeeds immediately, no retry loop entered).

This replaces the direct `.Font.Bold = True`-style assignments with calls to a small
`SetObjPropertyWithRetry` helper (added to `CF.bas`, using `CallByName` so one helper covers all
five properties). Per-property before/after logging is kept, now also reporting the retry count,
so we can see in the log whether/how often this is actually kicking in on the real report.

**Install:** replace the `CF` module with the version below (this supersedes the Round 1 CF.bas
— no need to keep the intermediate per-property-logging-only version). `DebugLog.bas` and
`MainModule.bas` are unchanged from the previous install.

**Please test this on the report that was hanging** and let me know: (a) whether it completes
without closing/hanging, and (b) if you can, send the new log tail so we can see the retry
counts — if it's regularly needing many retries, that's a sign the underlying "something is
busy" condition (possibly one of the other loaded add-ins) is worth chasing down directly rather
than just retried around.

## Round 3 findings (run 2026-08-15 20:30) — hang moved past Step-13/14, into untraced territory

With the retry wrapper in place, rule 1 of the conditional formats now succeeds cleanly on the
first try (`retries=0` for Color/Bold/Italic). Rule 2 this run hit the same busy error on
**`.FormatConditions.Add` itself** (not a property write) — a gap the retry wrapper didn't cover
yet, now closed below (`AddFormatConditionWithRetry`, also guards against touching `.Font`/
`.Interior` on a condition that never got created — that guard is what the previous version was
missing, and is why `Err=91` showed up right after the `Add` failure).

More importantly: the log runs cleanly all the way through Step-13 (conditional formats) *and*
Step-14 (label/value filters) — `"Applying pivot value and label filters Completed.."` is the
last line — and then Excel hangs with **no further log output**. That section of
`Generating_Pivots` (auto-expand/collapse, `nullDisplay`, and pivot protection via
`RestrictPivotTable`/`AllowPivotTable`/`Protecting_Sheet`) had zero logging before now, because
the working assumption was that the CF block was the only place this could happen. It clearly
isn't — `RestrictPivotTable`/`AllowPivotTable` do the exact same "several COM property writes
back-to-back in a loop over PivotFields" pattern that caused the original bug, just on
`PivotField.DragToPage/DragToRow/DragToColumn/DragToData/DragToHide` and
`PivotTable.EnableDrilldown/EnableFieldList/EnableFieldDialog/PivotCache.EnableRefresh` instead
of `FormatCondition.Font`/`.Interior`, and `ExpandCollapseDetails`/`ExpandCollapseFieldDetails`
repeatedly write `PivotField.ShowDetail` in a loop too.

The updated `MainModule.bas` below:
- Logs before/after the `autoExpandALL` block, the `nullDisplay` assignment, and the whole
  `pivotProtection` block (both the "protect" and "allow" branches, including that the "allow"
  branch previously logged nothing at all).
- Retry-wraps every one of those repeated property writes using the same
  `CF.SetObjPropertyWithRetry` helper (now `Public` in `CF.bas` so `MainModule` can call it too —
  `Sleep`/`RPC_E_SERVERCALL_RETRYLATER` don't need to be duplicated).
- Logs entry/exit and field name for `ExpandCollapseDetails`, `ExpandCollapseFieldDetails`,
  `RestrictPivotTable`, `AllowPivotTable`, `Protecting_Sheet`, `UnProtecting_Sheet`.

**Install:** replace both `CF` and `MainModule` with the versions below (`DebugLog.bas` is
unchanged). Reproduce the hang once more. If it's fixed, great — if it still hangs, the log will
now show exactly which of these newly-instrumented calls it's stuck on, the same way it did for
the CF block.

## Round 4 findings (run 2026-08-15 20:42 + `process.png`) — reordered expand-before-format

Log this run pinpointed the stall precisely inside `ExpandCollapseDetails`, triggered by
`autoExpandALL`/`noOfRowGroupsExpand=-1`:

```
ENTER ShowHide=True RowFields.Count=3
field=[SRG Order ID] ShowDetail=True -> Err=0 retries=0 success=True
<nothing further>
```

It's looping over 3 row fields (`SRG REGION`, `SRG STATE`, `SRG Order ID` — `SRG Order ID` is a
near-unique-per-record field). The first field processed succeeded instantly; it never returned
on the next one. Crucially, `process.png` (Task Manager) showed **Excel at 30.2% CPU, not 0%** —
so this is genuinely still computing, not a dead/deadlocked process waiting on a lock forever.

Why expanding row detail would be expensive here: by the time `autoExpandALL` ran (originally
positioned near the *end* of `Generating_Pivots`, after Step-13 conditional formatting and
Step-14 filters), the pivot already has `FormatCondition`s and `PivotFilter`s registered against
its *current* (mostly collapsed) row range. Forcing `ShowDetail=True` across all 3 row fields —
one of which is essentially one row per record — blows the row range open by a large factor, and
Excel has to reconcile/re-anchor every already-registered FormatCondition/filter against that
much bigger (and, per OISR-21412's original finding, likely non-contiguous/multi-area) range.
That reconciliation is the plausible source of the sustained ~30% CPU churn.

**Fix:** reorder `Generating_Pivots` so the auto-expand/collapse block runs *before* Step-13
(conditional formatting) and Step-14 (filters) instead of after. The pivot reaches its final row
shape first; formats and filters then get applied once, directly against that final shape, with
nothing needing to be migrated afterward. This is a pure reordering — no logic in either block
changed, no new fix needed in `CF.bas` for this. The `MainModule.bas` listing below has already
been updated with this reorder (search for `"Applying Auto Expand/Collapse.."` — it now runs
right after Step-12's number formatting and before Step-13's conditional formatting caption).

**Please test this reordered version** on the report that was hanging. If `ExpandCollapseDetails`
still grinds even when it's the *first* thing touching the pivot table (nothing yet registered
to reconcile), that would point away from the "reconciling existing formats" theory and toward
something more fundamentally expensive about expanding a field like `SRG Order ID` specifically
(e.g. genuinely too many distinct order IDs for the pivot engine to lay out quickly, in which
case the fix becomes bounding `noOfRowGroupsExpand` rather than reordering) — the log will make
that distinction clear either way.

## Round 5 findings — reproduced on a fresh file, root cause is `ScreenUpdating` vs. the Field List pane

The reorder fixed `ExpandCollapseDetails` completely (all 3 fields succeed instantly now). The
stall moved to `AllowPivotTable`, right after `EnableDrilldown=True` succeeds, before
`EnableFieldList`:

```
AllowPivotTable | ENTER PivotFields.Count=13
RetryHelper | AllowPivotTable EnableDrilldown=True -> Err=0 retries=0 success=True
<nothing further>
```

Ruled out, in order: other COM add-ins (none installed), OneDrive/SharePoint sync (file is on a
local, unsynced drive), and accumulated file cruft from earlier aborted runs (**reproduced
identically on a brand-new copy**, `DataReport_New.xlsm`, at the exact same call). A stall this
precisely reproducible regardless of file is not a race condition or environmental interference
— it's deterministic, which points at something structural in the code itself.

The structural culprit: `Speedup_Excel(True)` sets `Application.ScreenUpdating = False` for the
entire macro run (only restored at the very end of `Start_Pivot_Generation`). Every property that
stalled or errored in earlier rounds (`FormatConditions.Add`, `Font.Italic`, `PivotField.ShowDetail`)
is a plain internal flag/formula with no direct UI surface. `PivotTable.EnableFieldList`, by
contrast, is what attaches/shows the **PivotTable Field List task pane** — an actual UI element
that needs a real screen paint to appear. With `ScreenUpdating` off, Excel can end up waiting on
a repaint it's been told not to perform: a genuine, self-inflicted deadlock caused by our own
performance optimization, not an external race.

**Fix:** temporarily restore `Application.ScreenUpdating = True` around just the
`pivotProtection` block (`RestrictPivotTable`/`Protecting_Sheet` or `AllowPivotTable`), then set
it back to `False` immediately after, so the rest of the macro still gets the performance
benefit. Applied in the `MainModule.bas` listing below — search for
`"ScreenUpdating temporarily restored"`.

**Please test this** on the report that was hanging (a fresh copy is fine, or the same one — file
cruft is no longer a suspect). If `EnableFieldList` now succeeds and the macro completes, this
was it. If it *still* stalls at the same spot even with `ScreenUpdating` restored, the next thing
to try is skipping `CallByName` for this one property specifically (direct `PT.EnableFieldList =
True` assignment) in case late-bound reflection itself behaves differently for this particular
property — but the ScreenUpdating theory fits every piece of evidence collected so far, so it's
the strong favorite.

## Round 6 — ScreenUpdating fix ruled out; `EnableFieldList`/`EnableFieldDialog` removed instead

The `ScreenUpdating` restore ran (confirmed by `"ScreenUpdating temporarily restored for
pivotProtection block"` in the log) and made no difference — identical stall, same exact spot,
right after `EnableDrilldown=True` succeeds. That rules out the task-pane-repaint theory too.

At this point the stall has proven reproducible regardless of: other COM add-ins (none
installed), OneDrive/SharePoint sync (local unsynced drive), accumulated file cruft (fresh file,
same result), and `ScreenUpdating` state (on or off, same result). Six rounds of elimination on
an always-identical, always-reproducible symptom at the same exact property call is strong
evidence that something about `PivotTable.EnableFieldList` specifically and deterministically
blocks via `CallByName` in this environment, for reasons not worth chasing further given how
narrow its actual value is.

**Pragmatic fix:** `EnableFieldList` and `EnableFieldDialog` are no longer set at all in
`RestrictPivotTable`/`AllowPivotTable` (see the inline comments in the listing below). Both only
gate optional authoring-time UI (the Field List pane, the Field Settings dialog) — not report
data, layout, formatting, or the drill-down/drag interactions end users actually rely on, all of
which (`EnableDrilldown`, `PivotCache.EnableRefresh`, every `DragTo*` flag) already succeed fine
in every round's logs and remain unchanged. This sidesteps the confirmed blocking call entirely
rather than continuing to guess at its exact mechanism.

**Please test this final version.** If it completes, the macro is fixed end-to-end across all
the issues investigated in this document (OISR-21412's Areas-fragmentation fix, OISR-21439's
Font.Underline/retry-wrapped property writes, the auto-expand reorder, and now this). If the
report's users specifically need the Field List pane or Field Settings dialog available when
they open the generated file, that's a separate, much smaller ask we can look at once the hang
itself is confirmed gone — worth confirming that's even needed before spending more time on it.

## Round 7 — it was never `EnableFieldList` specifically; proactive yield added everywhere

With `EnableFieldList`/`EnableFieldDialog` removed, the log still stopped in the exact same
place - right after `EnableDrilldown=True` succeeds - because the *next* call in the sequence
(`PivotCache.EnableRefresh`) simply inherited the same block. That's the tell: this was never
about which property gets set second, only that *a* second COM property write follows a first
one that just succeeded. The same "1st write succeeds, 2nd write blocks" shape has now shown up
three times in three unrelated places (`CF.PivotCF`'s `Font.Bold`→`Font.Italic`,
`ExpandCollapseDetails`'s 2nd row field, and `AllowPivotTable`'s 2nd flag).

The retry wrapper only reacted *after* seeing an error, which is exactly why it couldn't help the
cases where the call never returns anything at all (no error to catch). The fix: both
`SetObjPropertyWithRetry` and `AddFormatConditionWithRetry` now call `DoEvents` + a short `Sleep
50` **after every successful write too**, not only inside the error-retry loop - giving Excel's
message loop a moment to settle before the very next COM call fires, proactively, instead of
reacting after the fact. This is a small, fixed cost per property write (a few dozen ms) applied
everywhere these helpers are already used (CF fonts/fills, `ShowDetail`, `EnableDrilldown`,
`PivotCache.EnableRefresh`, every `DragTo*` flag) — no per-callsite changes needed.

**Please test this version.** If it completes, this should be the fix for the whole class of
stalls seen across every round, not just this one call site.

## Round 8 — progress (pivot now visibly renders), still blocked; splitting GET from SET

Good news: the pivot table now visibly renders on screen mid-run (`New.png`) - it never did
before. That's the `ScreenUpdating` restore (Round 5) doing its job; the yield-on-success change
(Round 7) is confirmed active too (each successful write now takes ~60ms longer, matching the
added `Sleep 50`). But the log still stops in the same place, and this time `New.png`'s title bar
literally reads **"Progressing... (Not Responding)"** - Windows' own signal that the UI thread
hasn't serviced a message queue in a while.

That detail changes the read. The blocked call is:
```vba
Call CF.SetObjPropertyWithRetry(PT.PivotCache, "EnableRefresh", True, "AllowPivotTable PivotCache")
```
`PT.PivotCache` is a separate property **GET** (returns the PivotCache object), evaluated as an
argument *before* `SetObjPropertyWithRetry` is even entered - unlike every previous round, which
were all property *sets* on an object already in hand. If the GET itself is what blocks, that
explains why no log line ever appears for this call at all: the retry helper (and its `DoEvents`)
never starts running, because the block happens while evaluating the argument, before the
function call.

**Change:** in both `RestrictPivotTable` and `AllowPivotTable`, the `PT.PivotCache` GET is now a
separate, logged step (`Set PC = PT.PivotCache`, with `ThisWorkbook.PivotCaches.Count` logged
right before it), and `EnableRefresh` is set on that captured `PC` afterward. This will show,
for the first time, whether it's the GET or the SET that's actually blocking - and the
`PivotCaches.Count` log is worth checking too: if it's grown large (from PivotCaches never being
cleaned up across many aborted test runs on this same file across rounds 6-8), that's a second,
independent thing worth knowing regardless of which half turns out to block.

**Please test this version and send the new log + a screenshot if it stalls again.** Whichever
of "PivotCaches.Count=... BEFORE PT.PivotCache GET" / "AFTER PT.PivotCache GET" is the last line
tells us definitively which operation is the actual blocker.


Since the previous `CF.bas` fix didn't resolve the hang, this replaces guessing with **evidence**.
It adds a new logging module (`DebugLog`) plus instrumented versions of `CF.bas` and
`MainModule.bas` that write a timestamped line to a plain-text log **before and after every
meaningful step**, using `Open ... Close` around each write (not a buffered/held-open handle),
so every line is flushed to disk immediately — the log survives even if Excel has to be killed
from Task Manager mid-hang. Whatever the last line in the log is, that's where it's stuck.

## What you get

- **Every** `Progress.lblReport.Caption` step (already in the macro) is now also written to the
  log, so you get the full step timeline for free.
- Inside the Step-13 conditional-formatting block in `MainModule`, every row-label field,
  every data-label field, and every individual `conditionalFormat` rule is logged by name/index
  right before and right after the call into `CF.PivotCF`.
- Inside `CF.PivotCF` itself: field name, operator, data type, values, font flags on entry; the
  field's `DataRange.Areas.Count` and `Cells.Count` (this number alone will confirm or rule out
  the "thousands of disjoint Areas" theory from the last fix); each Area's address as it's
  processed; and — the most likely culprit — an explicit log line **immediately before** and
  **immediately after** the `.FormatConditions.Add(...)` call, since that's the one call in this
  whole path that hands control to Excel's own conditional-formatting engine and is the most
  plausible place for a genuine engine-side hang (as opposed to a VBA loop spinning).

## How to install

1. `Alt+F11` → VBA editor.
2. **Insert → Module**, name it exactly `DebugLog`, paste in the **DebugLog.bas** code below.
3. Open the existing **CF** module, `Ctrl+A` → `Delete`, paste in the **CF.bas** code below
   (this already includes both previous fixes — the per-Area loop and the `Underline` enum fix
   — plus the new logging; you don't need to re-apply the earlier fix separately).
4. Open the existing **MainModule**, `Ctrl+A` → `Delete`, paste in the **MainModule.bas** code
   below (only the two subs `Start_Pivot_Generation` and `Generating_Pivots` gained logging
   calls — every other sub in the module is byte-for-byte unchanged from the original).
5. Save as `.xlsm`.

`PivotFilters.bas` and `JsonConverter.bas` are unchanged — not reproduced here (see
`MACRO_DOCUMENTATION.md` for their contents if you need to confirm nothing changed).

## Where the log appears

`%LOCALAPPDATA%\ORBIT\Excel_Logs\PublisherReport_Debug.log`
(i.e. `C:\Users\<you>\AppData\Local\ORBIT\Excel_Logs\PublisherReport_Debug.log`)

The log is **cleared at the start of every run** (`DebugLog.ClearLog` is called once, right at
the top of `Start_Pivot_Generation`), so it only ever contains the current/last run — no need to
clean it up between attempts.

## What to do when it hangs again

1. Let it hang, then **open the log file directly** (Notepad, VS Code, anything) — don't wait for
   Excel, the file is already fully written up to the hang point since every write closes the
   file handle immediately.
2. Look at the **last line**. It'll be one of:
   - `"... BEFORE FormatConditions.Add"` with no matching `"AFTER FormatConditions.Add"` line
     after it → the hang is inside Excel's own conditional-formatting engine for that specific
     field/Area/formula. This is the most actionable case — send me that line (field name, area
     address, and the `formula=...` line just above it) and we'll target a fix at that exact
     rule.
   - A `"field=... area N/M ..."` line where `M` (total area count) is huge (hundreds/thousands)
     → confirms the DataRange-fragmentation theory; the fix becomes about avoiding
     per-Area application altogether (e.g. applying to one bounding contiguous range instead),
     rather than the per-Area loop.
   - Stuck mid-way through the row-labels or data-labels loop in `MainModule`, never reaching
     `CF.PivotCF` at all → the problem isn't in `CF.bas`, it's something upstream (e.g. reading
     `initialDict("conditionalFormat")`), and we'd look there instead.
3. Only after we know *exactly* which of these it is do we write the actual fix — that's the
   point of this pass.

Only kill Excel (Task Manager → End Task) once you've confirmed the log stopped advancing for a
minute or two — give it a little time in case it's just slow, not hung, on a very large report.

---

## `DebugLog.bas` (new module)

```vba
Attribute VB_Name = "DebugLog"
Option Explicit

' Set to False to silence logging once the hang is diagnosed and fixed.
Public Const DEBUG_LOG_ENABLED As Boolean = False

Private mStartTimer As Double
Private mLogPath As String

Private Sub EnsureLogPath()
    Dim FSO As Object
    Dim RootFolder As String
    Dim LogFolder As String

    If Len(mLogPath) > 0 Then Exit Sub

    RootFolder = Environ$("LOCALAPPDATA") & "\ORBIT"
    LogFolder = RootFolder & "\Excel_Logs"

    Set FSO = CreateObject("Scripting.FileSystemObject")
    If Not FSO.FolderExists(RootFolder) Then
        FSO.CreateFolder RootFolder
    End If
    If Not FSO.FolderExists(LogFolder) Then
        FSO.CreateFolder LogFolder
    End If

    mLogPath = LogFolder & "\PublisherReport_Debug.log"
    Set FSO = Nothing
End Sub

' Call once, right at the start of a run, to start a fresh log and reset the elapsed-time clock.
Public Sub ResetLog()
    Dim FileNum As Integer

    On Error Resume Next
    Call EnsureLogPath
    mStartTimer = Timer

    FileNum = FreeFile
    Open mLogPath For Output As #FileNum   ' Output = truncate/overwrite, i.e. clear the log
    Print #FileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | +0.000s | DebugLog | === RUN STARTED ==="
    Close #FileNum
    On Error GoTo 0
End Sub

' Call throughout the macro. Every call opens, appends one line, and closes the file
' immediately, so the line is guaranteed to be on disk even if Excel hangs/crashes right after.
Public Sub LogMsg(ByVal Source As String, ByVal Message As String)
    Dim FileNum As Integer
    Dim Elapsed As Double

    If Not DEBUG_LOG_ENABLED Then Exit Sub

    On Error Resume Next
    Call EnsureLogPath

    If mStartTimer = 0 Then mStartTimer = Timer
    Elapsed = Timer - mStartTimer
    If Elapsed < 0 Then Elapsed = Elapsed + 86400   ' guard against midnight rollover in Timer

    FileNum = FreeFile
    Open mLogPath For Append As #FileNum
    Print #FileNum, Format(Now, "yyyy-mm-dd hh:nn:ss") & " | +" & Format(Elapsed, "0.000") & "s | " & Source & " | " & Message
    Close #FileNum
    On Error GoTo 0
End Sub
```

---

## `CF.bas` (complete, with logging + the two previous fixes)

```vba
Attribute VB_Name = "CF"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Private Const RPC_E_SERVERCALL_RETRYLATER As Long = -2147417848

' Adds a formula-based FormatCondition to TargetRange, retrying on the same "server busy" COM
' error that .Add itself can also raise (confirmed in the field: it isn't only the Font/Interior
' property writes that can hit this). Returns Nothing if it never succeeds within MAX_RETRIES so
' the caller can skip font/fill formatting for that rule instead of touching an invalid object.
' Public so MainModule can reuse it too (RestrictPivotTable/AllowPivotTable/ExpandCollapse hit
' the identical class of COM "server busy" error on their own repeated property writes).
Public Function AddFormatConditionWithRetry(ByVal TargetRange As Range, ByVal FormulaStr As String, _
                                              ByVal LogLabel As String) As Object
    Const MAX_RETRIES As Integer = 10
    Dim RetryCount As Integer
    Dim FC As Object

    RetryCount = 0
    Do
        Err.Clear
        On Error Resume Next
        Set FC = TargetRange.FormatConditions.Add(Type:=xlExpression, Formula1:="=" & FormulaStr)
        On Error GoTo 0

        If Err.Number = 0 Then
            ' Yield even on success - see matching comment in SetObjPropertyWithRetry.
            DoEvents
            Sleep 50
            Exit Do
        End If
        If Err.Number <> RPC_E_SERVERCALL_RETRYLATER Then Exit Do

        RetryCount = RetryCount + 1
        DoEvents
        Sleep 200
    Loop While RetryCount < MAX_RETRIES

    Call DebugLog.LogMsg("RetryHelper", LogLabel & " FormatConditions.Add -> Err=" & Err.Number & " retries=" & RetryCount)
    Set AddFormatConditionWithRetry = FC
End Function

' Sets TargetObj.PropName = PropValue via CallByName (so this one helper covers Font.Color,
' Font.Bold, Font.Italic, Font.Underline, Interior.Color, PivotField.ShowDetail/DragTo*, etc.
' alike - anything that's a simple property Let). On the COM "server busy" error that Excel
' intermittently raises when properties are set back-to-back on the same object, this yields to
' the message loop (DoEvents) and waits briefly before retrying, instead of either silently
' giving up (the original bug) or hanging forever. Public so MainModule can reuse it too.
Public Function SetObjPropertyWithRetry(ByVal TargetObj As Object, ByVal PropName As String, _
                                          ByVal PropValue As Variant, ByVal LogLabel As String) As Boolean
    Const MAX_RETRIES As Integer = 10
    Dim RetryCount As Integer

    RetryCount = 0
    Do
        Err.Clear
        On Error Resume Next
        CallByName TargetObj, PropName, VbLet, PropValue
        On Error GoTo 0

        If Err.Number = 0 Then
            ' Yield even on success. Confirmed in the field: it is not any particular property
            ' name that blocks - it is firing a second COM property write immediately after a
            ' first one succeeds (seen with Font.Bold->Font.Italic, ExpandCollapseDetails's 2nd
            ' row field, and EnableDrilldown->whatever came next). Giving Excel's message loop a
            ' moment to settle after every successful write, not only after a failed one, is the
            ' proactive version of the same fix.
            DoEvents
            Sleep 50
            SetObjPropertyWithRetry = True
            Exit Do
        End If

        If Err.Number <> RPC_E_SERVERCALL_RETRYLATER Then
            SetObjPropertyWithRetry = False
            Exit Do
        End If

        RetryCount = RetryCount + 1
        DoEvents
        Sleep 200
    Loop While RetryCount < MAX_RETRIES

    Call DebugLog.LogMsg("RetryHelper", LogLabel & " " & PropName & "=" & PropValue & _
        " -> Err=" & Err.Number & " retries=" & RetryCount & " success=" & SetObjPropertyWithRetry)
End Function

Public Sub PivotCF(ByVal PT As PivotTable, ByVal JsonDict As Dictionary, ByVal PFFieldName As String, ByVal IsPivotRowField As Boolean)
    Dim pf As PivotField
    Dim rng As Range
    Dim ar As Range
    Dim FormulaStr As String
    Dim CFOperator As String
    Dim CFFontBold As String
    Dim CFFontItalic As String
    Dim CFFontUnderline As String
    Dim CFFontFillColor As String
    Dim CFFontColor As String
    Dim CFFontForeColor As String
    Dim CFValue1 As Variant
    Dim CFValue2 As Variant
    Dim RngAdd As String
    Dim FormatRng As Range
    Dim FormatStr As String
    Dim StrAddr As String
    Dim DTType As String
    Dim DataDef As String
    Dim CFColumnName As String
    Dim pfCol As PivotField
    Dim pfColRange As Range
    Dim AreaIndex As Long
    Dim AreaCount As Long
    Dim LogLabel As String
    Dim FC As Object

    Call DebugLog.LogMsg("CF.PivotCF", "ENTER field=[" & PFFieldName & "] isRowField=" & IsPivotRowField)

    If JsonDict.Exists("operator") Then
       CFOperator = JsonDict("operator")
    Else
       CFOperator = ""
    End If

    If Len(CFOperator) = 0 Or CFOperator = "" Then
        Call DebugLog.LogMsg("CF.PivotCF", "EXIT field=[" & PFFieldName & "] - no operator, nothing to do")
        Exit Sub
    End If

    CFColumnName = JsonDict("columnName")
    DTType = JsonDict("dataType")
    CFFontBold = JsonDict("fontBold")
    CFFontItalic = JsonDict("fontItalic")
    CFFontUnderline = JsonDict("fontUnderline")
    CFFontForeColor = JsonDict("fontForeColor")
    CFFontFillColor = JsonDict("cellFillColor")

    Set pf = PT.PivotFields(PFFieldName)

    CFValue1 = JsonDict("value1")
    CFValue2 = JsonDict("value2")

    Call DebugLog.LogMsg("CF.PivotCF", "field=[" & PFFieldName & "] operator=" & CFOperator & " dataType=" & DTType & _
        " value1=" & CFValue1 & " value2=" & CFValue2 & " bold=" & CFFontBold & " italic=" & CFFontItalic & _
        " underline=" & CFFontUnderline & " fillColor=" & CFFontFillColor & " foreColor=" & CFFontForeColor)

    On Error Resume Next
    Set pfCol = PT.PivotFields(fieldMap(CFColumnName))
    If Not pfCol Is Nothing Then
       Set pfColRange = pfCol.DataRange.Cells(1, 1)
    Else
       Set pfCol = Nothing
    End If
    On Error GoTo 0

    If UCase(Trim(DTType)) = "DATE" Or UCase(Trim(DTType)) = "TIME" Or UCase(Trim(DTType)) = "DATETIME" Then
       DataDef = "DATE"
    ElseIf UCase(Trim(DTType)) = "STRING" Or UCase(Trim(DTType)) = "CHAR" Or UCase(Trim(DTType)) = "VARCHAR" Or UCase(Trim(DTType)) = "NVARCHAR" Then
       DataDef = "STRING"
    ElseIf UCase(Trim(DTType)) = "INTEGER" Or UCase(Trim(DTType)) = "DECIMAL" Or UCase(Trim(DTType)) = "NUMBER" Or UCase(Trim(DTType)) = "NUMERIC" Or _
           UCase(Trim(DTType)) = "INT" Or UCase(Trim(DTType)) = "DOUBLE" Then
       DataDef = "NUMBER"
    End If

    AreaCount = pf.DataRange.Areas.Count
    Call DebugLog.LogMsg("CF.PivotCF", "field=[" & PFFieldName & "] pf.DataRange.Cells.Count=" & pf.DataRange.Cells.Count & _
        " pf.DataRange.Areas.Count=" & AreaCount)

    AreaIndex = 0
    For Each ar In pf.DataRange.Areas
    AreaIndex = AreaIndex + 1
    Call DebugLog.LogMsg("CF.PivotCF", "field=[" & PFFieldName & "] area " & AreaIndex & "/" & AreaCount & " address=" & ar.Address & " cells=" & ar.Cells.Count)
    With ar
        Set rng = ar.Cells(1, 1)
        If Not pfColRange Is Nothing Then
           RngAdd = Cells(rng.Row, pfColRange.Column).Address(False, False)
        Else
           RngAdd = rng.Address(False, False)
        End If

        StrAddr = Range("A" & rng.Row).Address & ":" & rng.Address
        Set FormatRng = Range(StrAddr)

        FormatStr = "ISERROR(SEARCH(""Total""," & FormatRng.Address(False, False) & "))"

        Select Case UCase(CFOperator)
            Case "XLBETWEEN"
                If IsPivotRowField = True Then
                   If DataDef = "DATE" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", " & RngAdd & ">=" & Chr(34) & CFValue1 & Chr(34) & ", " & RngAdd & "<=" & Chr(34) & CFValue2 & Chr(34) & ", " & FormatStr & ")"
                   ElseIf DataDef = "STRING" Then
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", AND(" & RngAdd & ">=" & Chr(34) & CFValue1 & Chr(34) & "," & _
                                RngAdd & "<=" & Chr(34) & CFValue2 & Chr(34) & "), " & FormatStr & ")"
                   ElseIf DataDef = "NUMBER" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(" & RngAdd & ")," & RngAdd & " >= " & CFValue1 & ", IF(ISNUMBER(" & RngAdd & ")," & RngAdd & ",0) <= " & CFValue2 & ", " & FormatStr & ")"
                   End If
                Else
                   FormulaStr = "AND(" & RngAdd & "<>""""" & ", AND(" & RngAdd & ">=" & CFValue1 & "," & RngAdd & "<=" & CFValue2 & "))"
                End If
            Case "XLNOTBETWEEN"
                If IsPivotRowField = True Then
                    If DataDef = "NUMBER" Then
                        FormulaStr = "AND(" & RngAdd & "<>"""", OR(ISNUMBER(" & RngAdd & ")," & RngAdd & " < " & CFValue1 & ", ISNUMBER(" & RngAdd & ")," & RngAdd & " > " & CFValue2 & "), " & FormatStr & ")"
                    ElseIf DataDef = "STRING" Then
                        FormulaStr = "AND(" & RngAdd & "<>"""", OR(" & RngAdd & "<" & Chr(34) & CFValue1 & Chr(34) & ", " & RngAdd & ">" & Chr(34) & CFValue2 & Chr(34) & "), " & FormatStr & ")"
                    End If
                Else
                    FormulaStr = "AND(" & RngAdd & "<>"""", OR(" & RngAdd & "<" & CFValue1 & ", " & RngAdd & ">" & CFValue2 & "))"
                End If
            Case "XLEQUAL"
                If IsPivotRowField = True Then
                   If DataDef = "DATE" Then
                      CFValue1 = DateSerial(Year(CFValue1), Month(CFValue1), Day(CFValue1))
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "=DATEVALUE(" & Chr(34) & CFValue1 & Chr(34) & "), " & FormatStr & ")"
                   ElseIf DataDef = "STRING" Then
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "=" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
                   ElseIf DataDef = "NUMBER" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(" & RngAdd & ")," & RngAdd & " = " & CFValue1 & ", " & FormatStr & ")"
                   End If
                Else
                   FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "=" & CFValue1 & ")"
                End If
            Case "XLNOTEQUAL"
                If IsPivotRowField = True Then
                   If DataDef = "DATE" Then
                      CFValue1 = DateSerial(Year(CFValue1), Month(CFValue1), Day(CFValue1))
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<>DATEVALUE(" & Chr(34) & CFValue1 & Chr(34) & "), " & FormatStr & ")"
                   ElseIf DataDef = "STRING" Then
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<>" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
                   ElseIf DataDef = "NUMBER" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(" & RngAdd & ")," & RngAdd & " <> " & CFValue1 & ", " & FormatStr & ")"
                   End If
                Else
                   FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<>" & CFValue1 & ")"
                End If
            Case "XLGREATER"
                If IsPivotRowField = True Then
                   If DataDef = "DATE" Then
                      CFValue1 = DateSerial(Year(CFValue1), Month(CFValue1), Day(CFValue1))
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & ">DATEVALUE(" & Chr(34) & CFValue1 & Chr(34) & "), " & FormatStr & ")"
                   ElseIf DataDef = "STRING" Then
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & ">" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
                   ElseIf DataDef = "NUMBER" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(" & RngAdd & ")," & RngAdd & " > " & CFValue1 & ", " & FormatStr & ")"
                   End If
                Else
                   FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & ">" & CFValue1 & ")"
                End If
            Case "XLGREATEREQUAL"
                If IsPivotRowField = True Then
                   If DataDef = "DATE" Then
                      CFValue1 = DateSerial(Year(CFValue1), Month(CFValue1), Day(CFValue1))
                      FormulaStr = "AND(" & RngAdd & "<> """"" & ", " & RngAdd & ">= DATEVALUE(" & Chr(34) & CFValue1 & Chr(34) & "), " & FormatStr & ")"
                   ElseIf DataDef = "STRING" Then
                      FormulaStr = "AND(" & RngAdd & "<> """"" & ", " & RngAdd & ">= " & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
                   ElseIf DataDef = "NUMBER" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(" & RngAdd & ")," & RngAdd & " >= " & CFValue1 & ", " & FormatStr & ")"
                   End If
                Else
                   FormulaStr = "AND(" & RngAdd & "<> """"" & ", " & RngAdd & ">= " & CFValue1 & ")"
                End If
            Case "XLLESS"
                If IsPivotRowField = True Then
                   If DataDef = "DATE" Then
                      CFValue1 = DateSerial(Year(CFValue1), Month(CFValue1), Day(CFValue1))
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<DATEVALUE(" & Chr(34) & CFValue1 & Chr(34) & "), " & FormatStr & ")"
                   ElseIf DataDef = "STRING" Then
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
                   ElseIf DataDef = "NUMBER" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(" & RngAdd & ")," & RngAdd & " < " & CFValue1 & ", " & FormatStr & ")"
                   End If
                Else
                   FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<" & CFValue1 & ")"
                End If
            Case "XLLESSEQUAL"
                If IsPivotRowField = True Then
                   If DataDef = "DATE" Then
                      CFValue1 = DateSerial(Year(CFValue1), Month(CFValue1), Day(CFValue1))
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<=DATEVALUE(" & Chr(34) & CFValue1 & Chr(34) & "), " & FormatStr & ")"
                   ElseIf DataDef = "STRING" Then
                      FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<=" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
                   ElseIf DataDef = "NUMBER" Then
                      FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(" & RngAdd & ")," & RngAdd & " <= " & CFValue1 & ", " & FormatStr & ")"
                   End If
                Else
                   FormulaStr = "AND(" & RngAdd & "<>""""" & ", " & RngAdd & "<=" & CFValue1 & ")"
                End If
            Case "XLCONTAINS"
                FormulaStr = "AND(" & RngAdd & "<>"""", ISNUMBER(SEARCH(" & Chr(34) & CFValue1 & Chr(34) & "," & RngAdd & ")), " & FormatStr & ")"
            Case "XLDOESNOTCONTAIN"
                FormulaStr = "AND(" & RngAdd & "<>"""", ISERROR(SEARCH(" & Chr(34) & CFValue1 & Chr(34) & "," & RngAdd & ")), " & FormatStr & ")"
            Case "XLBEGINSWITH"
                FormulaStr = "AND(" & RngAdd & "<>"""", LEFT(" & RngAdd & ",LEN(" & Chr(34) & CFValue1 & Chr(34) & "))=" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
            Case "XLDOESNOTBEGINSWITH"
                FormulaStr = "AND(" & RngAdd & "<>"""", RIGHT(" & RngAdd & ",LEN(" & Chr(34) & CFValue1 & Chr(34) & "))=" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
            Case "XLENDSWITH"
                FormulaStr = "AND(" & RngAdd & "<>"""", RIGHT(" & RngAdd & ",LEN(" & Chr(34) & CFValue1 & Chr(34) & "))=" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
            Case "XLDOESNOTENDSWITH"
                FormulaStr = "AND(" & RngAdd & "<>"""", RIGHT(" & RngAdd & ",LEN(" & Chr(34) & CFValue1 & Chr(34) & "))<>" & Chr(34) & CFValue1 & Chr(34) & ", " & FormatStr & ")"
            Case "XLISEMPTY"
                FormulaStr = "OR(" & RngAdd & "="""", LEN(" & RngAdd & ")=0, " & RngAdd & "=""(blank)"")"
            Case "XLISNOTEMPTY"
                FormulaStr = "AND(" & RngAdd & "<>"""", LEN(" & RngAdd & ")>0, " & FormatStr & ")"
        End Select

        Call DebugLog.LogMsg("CF.PivotCF", "field=[" & PFFieldName & "] area " & AreaIndex & " formula==" & FormulaStr)

        On Error Resume Next
        LogLabel = "field=[" & PFFieldName & "] area " & AreaIndex
        Call DebugLog.LogMsg("CF.PivotCF", LogLabel & " BEFORE FormatConditions.Add")

        Set FC = AddFormatConditionWithRetry(ar, FormulaStr, LogLabel)

        If FC Is Nothing Then
            Call DebugLog.LogMsg("CF.PivotCF", LogLabel & " FormatConditions.Add never succeeded - skipping font/fill for this rule")
        Else
            If Len(CFFontForeColor) > 0 And CFFontForeColor <> "" Then
                Call SetObjPropertyWithRetry(FC.Font, "Color", GetLongFromHex(CFFontForeColor), LogLabel)
            End If

            If Len(CFFontFillColor) > 0 And CFFontFillColor <> "" Then
                Call SetObjPropertyWithRetry(FC.Interior, "Color", GetLongFromHex(CFFontFillColor), LogLabel)
            End If

            If Len(CFFontBold) > 0 And CFFontBold <> "" And CFFontBold = "Y" Then
                Call SetObjPropertyWithRetry(FC.Font, "Bold", True, LogLabel)
            End If

            If Len(CFFontItalic) > 0 And CFFontItalic <> "" And CFFontItalic = "Y" Then
                Call SetObjPropertyWithRetry(FC.Font, "Italic", True, LogLabel)
            End If

            If Len(CFFontUnderline) > 0 And CFFontUnderline <> "" And CFFontUnderline = "Y" Then
                Call SetObjPropertyWithRetry(FC.Font, "Underline", CLng(xlUnderlineStyleSingle), LogLabel)
            End If
        End If
        Call DebugLog.LogMsg("CF.PivotCF", LogLabel & " font/fill/style applied (Err=" & Err.Number & ")")
        On Error GoTo 0
    End With
    Next ar

    On Error Resume Next
    Set rng = Nothing
    Set pf = Nothing
    Set FormatRng = Nothing
    On Error GoTo 0

    Call DebugLog.LogMsg("CF.PivotCF", "EXIT field=[" & PFFieldName & "] - completed " & AreaIndex & " area(s)")
End Sub
Private Function GetLongFromHex(hexColor As String) As Long

Dim R As String
Dim G As String
Dim B As String

hexColor = VBA.Replace(hexColor, "#", "")
hexColor = VBA.Right$("000000" & hexColor, 6)

R = Left(hexColor, 2)
G = Mid(hexColor, 3, 2)
B = Right(hexColor, 2)

GetLongFromHex = Application.WorksheetFunction.Hex2Dec(B & G & R)
End Function
```

---

## `MainModule.bas` (complete, with logging added around every step and deep inside Step-13)

```vba
Attribute VB_Name = "MainModule"
Option Explicit

#If VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
#End If

Public fieldMap As Scripting.Dictionary
Private Function ValidFileName(ByVal FileName As String, ByRef chrs As String) As Boolean
'PURPOSE: Determine If A Given Excel File Name Is Valid
'SOURCE: www.TheSpreadsheetGuru.com/the-code-vault
'AUTHOR: Jon Peltier

Const sBadChar As String = "\/:*?<>|[]"""
Dim i As Long

'Assume valid unless it isn't
  ValidFileName = True

'Loop through each "Bad Character" and test for an instance
  For i = 1 To Len(sBadChar)
    If InStr(FileName, Mid$(sBadChar, i, 1)) > 0 Then
      ValidFileName = False 'Invalid
      chrs = Mid(sBadChar, i, 1)
      Exit For
    End If
  Next

End Function
Private Sub Speedup_Excel(ByVal BoolVal As Boolean)
With Application
   If BoolVal = True Then
      .ScreenUpdating = False
      .DisplayAlerts = False
      .Calculation = xlCalculationManual
      .EnableEvents = False
   Else
      .Calculation = xlCalculationAutomatic
      .EnableEvents = True
      .ScreenUpdating = True
      .DisplayAlerts = True
   End If
End With
End Sub
Private Sub InitUFProgressBarBar()
    
    With Progress
        .Bar.Width = 0
        .Text.Caption = "0% Complete"
        .Show vbModeless
    End With
End Sub
Private Sub ProgressBar_Chart(ByVal iStep As Long)
    
    DoEvents
    
    Dim CurrentUFProgressBar As Double
    Dim UFProgressBarPercentage As Double
    Dim BarWidth As Long

    CurrentUFProgressBar = iStep / 100
    BarWidth = Progress.Border.Width * CurrentUFProgressBar
    UFProgressBarPercentage = Round(CurrentUFProgressBar * 100, 0)
    
    Progress.Bar.Width = BarWidth
    Progress.Text.Caption = UFProgressBarPercentage & "% Complete"
    
    DoEvents
    
    
End Sub
Private Function ORB_SheetExists(ByVal WorksheetName As String) As Boolean
    On Error Resume Next
    ORB_SheetExists = (Sheets(WorksheetName).Name <> "")
    On Error GoTo 0
End Function
Public Sub Start_Pivot_Generation()

    Dim isValidFile As Boolean
    Dim FileSplits As Variant
    Dim InvalidChar As String
    Dim PivotRange As Range
    Dim ConfigSheet As Worksheet
    Dim RawSheet As Worksheet
    Dim PivotSheet As Worksheet
    Dim PCache As PivotCache, PT As PivotTable
    Dim lastRow As Long, lastColumn As Long
    
    Call DebugLog.ResetLog
    Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "=== Start_Pivot_Generation begin ===")
    
    Call Speedup_Excel(True)
    
    On Error Resume Next
    
    DoEvents
    Call InitUFProgressBarBar
    Progress.lblReport.Caption = "Validating File Name"
    Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
    Call ProgressBar_Chart(7)
    DoEvents
'Step-1: Checking if file name is valid.
    If InStr(1, ThisWorkbook.FullName, "\") = 0 Then
       GoTo Exit_OnHere
    End If
    
    FileSplits = Split(ThisWorkbook.FullName, "\")
    
    If UBound(FileSplits) = 0 Then
       GoTo Exit_OnHere
    End If
    
    isValidFile = ValidFileName(FileSplits(UBound(FileSplits)), InvalidChar)
    If isValidFile = False Then
       DoEvents
       Progress.lblReport.Caption = "Validation Failed!!!"
       Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
       DoEvents
       Unload Progress
       MsgBox "The file name has invalid character(s)." & vbCrLf & "The character is " & Chr(34) & InvalidChar & Chr(34), vbInformation, "Pivot"
       Call Speedup_Excel(False)
       Exit Sub
    End If
    
Exit_OnHere:
    On Error GoTo 0
    DoEvents
    Progress.lblReport.Caption = "Checking for all sheets and data"
    Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
    Call ProgressBar_Chart(14)
    DoEvents
'Step-2: Checking if all sheets exists,raw data sheet has data for creating pivot table.
    If ORB_SheetExists("RawData") = False Then
       DoEvents
       Progress.lblReport.Caption = "Missing Raw Data Sheet!!!"
       Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
       DoEvents
       Unload Progress
       Call Speedup_Excel(False)
       Exit Sub
    Else
        Set RawSheet = ThisWorkbook.Worksheets("RawData")
        
        On Error Resume Next
          lastRow = RawSheet.Cells.Find(what:="*", After:=RawSheet.Range("A1"), lookat:=xlPart, LookIn:=xlFormulas, _
          searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
        On Error GoTo 0
        
        On Error Resume Next
          lastColumn = RawSheet.Cells.Find(what:="*", After:=RawSheet.Range("A1"), lookat:=xlPart, LookIn:=xlFormulas, _
          searchorder:=xlByColumns, SearchDirection:=xlPrevious, MatchCase:=False).Column
        On Error GoTo 0
        
        If (lastRow <= 1 Or lastColumn <= 0) Or RawSheet.Range("A1").Value = "${TableHeader}" Then
           Unload Progress
           MsgBox "Raw data sheet is empty!!!", vbInformation, ThisWorkbook.Name
           Call Speedup_Excel(False)
           Exit Sub
        End If
        
        Set PivotRange = RawSheet.Range(RawSheet.Cells(1, 1), RawSheet.Cells(lastRow, lastColumn))
        
    End If
    
    If ORB_SheetExists("PivotOrbConfig") = False Then
       DoEvents
       Progress.lblReport.Caption = "Missing Configuration Sheet!!!"
       Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
       DoEvents
       Unload Progress
       Call Speedup_Excel(False)
       Exit Sub
    Else
        Set ConfigSheet = ThisWorkbook.Worksheets("PivotOrbConfig")
        If ConfigSheet.Range("A1").Value = "" Or Len(ConfigSheet.Range("A1").Value) = 0 Then
           DoEvents
           Progress.lblReport.Caption = "No Meta Information to Proceed!!!"
           Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
           DoEvents
           Unload Progress
           Call Speedup_Excel(False)
           Exit Sub
        End If
    End If
    
    If ORB_SheetExists("ReportOutput") = False Then
       DoEvents
       Progress.lblReport.Caption = "Missing ReportOutput Sheet!!!"
       Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
       DoEvents
       Unload Progress
       Call Speedup_Excel(False)
       Exit Sub
    Else
        Set PivotSheet = ThisWorkbook.Worksheets("ReportOutput")
    End If
    DoEvents
    Progress.lblReport.Caption = "Sheet validation completed"
    Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
    DoEvents
    
'Checking if macro has already executed not.
DoEvents
Progress.lblReport.Caption = "Checking Macro Execution Status"
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
DoEvents
If UCase(ConfigSheet.Range("A3").Value) = "Y" Then
   DoEvents
   Progress.lblReport.Caption = "Macro Already Executed!!!"
   Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
   DoEvents
   Unload Progress
   MsgBox "Macro Already Executed!!!" & vbCrLf & _
          "Clear " & Chr(34) & "Y" & Chr(34) & _
          " from cell A3 of " & Chr(34) & "PivotOrbConfig" & Chr(34) & " worksheet, which is hidden.", vbInformation, ThisWorkbook.Name
   Call Speedup_Excel(False)
   Exit Sub
End If

'Step-3: Parsing JSON string
Dim Dict As New Scripting.Dictionary
On Error Resume Next
Set Dict = JsonConverter.ParseJSON(ConfigSheet.Range("A1").Value)
If Dict Is Nothing Then
   DoEvents
   Progress.lblReport.Caption = "JSON Validation Failed!!!"
   Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
   DoEvents
   Unload Progress
   MsgBox "Invalid JSON...", vbCritical, "Orbit BI"
   Call Speedup_Excel(False)
   Exit Sub
End If
On Error GoTo 0
'Step-3: Formatting raw data sheet.

DoEvents
Progress.lblReport.Caption = "Formatting Raw Data Sheet.."
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
Call ProgressBar_Chart(21)
DoEvents
If Dict.Exists("rawData") Then
   Call RawSheetFormat(RawSheet, Dict, PivotRange)
End If
DoEvents
Progress.lblReport.Caption = "Formatting Completed.."
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
DoEvents

'Step-4: Starting pivot generation
DoEvents
Progress.lblReport.Caption = "Pivot Generation Started.."
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
Call ProgressBar_Chart(28)
DoEvents
Call Generating_Pivots(Dict, PivotSheet, PivotRange)
DoEvents
Progress.lblReport.Caption = "Pivot Generation Completed.."
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
DoEvents



DoEvents
Progress.lblReport.Caption = "Other Miscellaneous.."
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
Call ProgressBar_Chart(84)
DoEvents
'Step-14: Saving file as XLSX from XLSM and deleting XLSM file if delete flag is true
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "BEFORE saveXlsmToXlsx check")
If Dict.Exists("saveXlsmToXlsx") And LCase(TypeName(Dict("saveXlsmToXlsx"))) = "string" And UCase(Dict("saveXlsmToXlsx")) = "Y" Then
   Dim FileConv As String
   Dim FileToDelete As String
   Dim FSO
   
   Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "saveXlsmToXlsx=Y, entering SaveAs block")
   FileToDelete = ActiveWorkbook.FullName
   FileConv = Left(ActiveWorkbook.FullName, InStrRev(ActiveWorkbook.FullName, ".") - 1)
   
   Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "BEFORE ActiveWorkbook.SaveAs FileConv=[" & FileConv & "]")
   ActiveWorkbook.SaveAs FileName:=FileConv, FileFormat _
                          :=xlOpenXMLWorkbook, CreateBackup:=False
   Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "AFTER ActiveWorkbook.SaveAs (Err=" & Err.Number & ")")
   If Dict.Exists("deleteXlsmFile") And LCase(TypeName(Dict("deleteXlsmFile"))) = "string" And UCase(Dict("deleteXlsmFile")) = "Y" Then
      Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "deleteXlsmFile=Y, BEFORE FSO.Deletefile [" & FileToDelete & "]")
      Set FSO = CreateObject("Scripting.FileSystemObject")
      FSO.Deletefile FileToDelete, True
      Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "AFTER FSO.Deletefile (Err=" & Err.Number & ")")
      Set FSO = Nothing
   End If
Else
   Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "saveXlsmToXlsx not Y, skipping SaveAs block")
End If
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "AFTER saveXlsmToXlsx check")

DoEvents
Progress.lblReport.Caption = "All Tasks Completed.."
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", Progress.lblReport.Caption)
Call ProgressBar_Chart(100)
Unload Progress
DoEvents

Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "BEFORE ConfigSheet.Range(A3) write")
ConfigSheet.Range("A3").Value = "Y"
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "BEFORE Application.GoTo")
Application.GoTo reference:=PivotSheet.Range("A1"), Scroll:=True
Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "AFTER Application.GoTo")

Set Dict = Nothing
Set PivotRange = Nothing
Set PCache = Nothing
Set PT = Nothing
Set ConfigSheet = Nothing
Set RawSheet = Nothing
Set PivotSheet = Nothing

Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "BEFORE final Speedup_Excel(False)")
Call Speedup_Excel(False)

Call DebugLog.LogMsg("MainModule.Start_Pivot_Generation", "=== Start_Pivot_Generation end ===")

End Sub
Private Sub RawSheetFormat(ByVal sht As Worksheet, ByVal JsonDict As Dictionary, ByVal rng As Range)
Dim KeyStr As Variant
Dim initialCollection  As Collection
Dim initialDict As Dictionary
Dim i As Integer
Dim key As Variant

On Error Resume Next
If LCase(TypeName(JsonDict("rawData"))) = "collection" Or LCase(TypeName(JsonDict("rawData"))) = "dictionary" Then
   Set initialCollection = JsonDict("rawData")
   If initialCollection.Count > 0 Then
      For i = 1 To initialCollection.Count
         If LCase(TypeName(initialCollection(i))) = "dictionary" Then
            Set initialDict = initialCollection(i)
            If initialDict.Count > 0 Then
               If initialDict.Exists("columnName") And initialDict.Exists("numberFormat") And _
                  initialDict.Exists("dataType") Then
                  Call Format_RawSheet(sht, rng, initialDict("columnName"), initialDict("numberFormat"), initialDict("dataType"))
                  'Debug.Print initialDict("columnName") & " = " & initialDict("dataType")
               End If
               'Debug.Print key & " = " & TypeName(initialDict(key))
            End If
         End If
      Next
   End If
End If
Set initialDict = Nothing
Set initialCollection = Nothing
On Error GoTo 0
End Sub
Private Sub Generating_Pivots(ByVal JsonDict As Dictionary, ByVal PivotSheet As Worksheet, ByVal rng As Range)
Dim KeyStr As Variant
Dim PCache As PivotCache, PT As PivotTable
Dim lastRow As Long, lastColumn As Long
Dim DummyRng As Range
Dim FieldName As String, SortName As String, numberFormat As String, showNegative As String, applyFormat As String
Dim i As Integer, j As Integer
Dim initialCollection  As Collection
Dim initialDict As Dictionary
Dim initialDict1 As Dictionary
Dim DictKey As Variant
Dim FormatStr As String
Dim DataRange As Range

Call DebugLog.LogMsg("MainModule.Generating_Pivots", "=== Generating_Pivots begin ===")

On Error Resume Next
PivotSheet.Activate

Set fieldMap = New Scripting.Dictionary

Set PCache = ThisWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=rng.Address(False, False, xlR1C1, xlExternal))
If PivotSheet.PivotTables.Count > 0 Then
   If PivotSheet.PivotTables(1).Name = "OrbitPivot" Then
      If CheckIfSheetIsProtected(PivotSheet) = True Then
         Call UnProtecting_Sheet(PivotSheet)
      End If
      Call AllowPivotTable(PivotSheet.PivotTables(1))
      PivotSheet.PivotTables(1).TableRange2.Clear
   End If
End If

lastRow = 0

On Error Resume Next
    lastRow = PivotSheet.Cells.Find(what:="*", After:=PivotSheet.Range("A1"), lookat:=xlPart, LookIn:=xlFormulas, _
      searchorder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row
On Error GoTo 0
If lastRow = 0 Then
   lastRow = 1
Else
   lastRow = lastRow + 2
End If

'PivotSheet.Cells.Clear.

Set DummyRng = PivotSheet.Range("A" & lastRow)
DummyRng.Select
Range(Selection, Selection.End(xlToRight)).Select
Range(Selection, Selection.End(xlDown)).Select
Selection.Clear

Set PT = PivotSheet.PivotTables.Add(PivotCache:=PCache, TableDestination:=DummyRng, TableName:="OrbitPivot")
PT.DisplayErrorString = True
DoEvents
Progress.lblReport.Caption = "Applying Pivot Layout.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(35)
DoEvents
'Step-5: Pivotlayout.
If JsonDict.Exists("pivotLayout") And LCase(TypeName(JsonDict("pivotLayout"))) = "string" Then
    If UCase(JsonDict("pivotLayout")) = "COMPACT" Then
        PT.RowAxisLayout xlCompactRow
    ElseIf UCase(JsonDict("pivotLayout")) = "OUTLINE" Then
        PT.RowAxisLayout xlOutlineRow
    ElseIf UCase(JsonDict("pivotLayout")) = "TABULAR" Then
        PT.RowAxisLayout xlTabularRow
    End If
End If
DoEvents
Progress.lblReport.Caption = "Layout Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

DoEvents
Progress.lblReport.Caption = "Adding Page Filters.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(42)
DoEvents
'Step-6: Pivot page filters.
If JsonDict.Exists("pageFilters") And LCase(TypeName(JsonDict("pageFilters"))) = "collection" Then
   Set initialCollection = JsonDict("pageFilters")
   If initialCollection.Count > 0 Then
      For i = initialCollection.Count To 1 Step -1
         FieldName = initialCollection(i)
         fieldMap.Add FieldName, FieldName
         With PT.PivotFields(FieldName)
              .Orientation = xlPageField
         End With
      Next
   End If
End If
DoEvents
Progress.lblReport.Caption = "Page Filters Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

DoEvents
Progress.lblReport.Caption = "Adding Row Labels.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(49)
DoEvents
'Step-7: Pivot row labels.

If JsonDict.Exists("rowLabels") And LCase(TypeName(JsonDict("rowLabels"))) = "collection" Then
   Set initialCollection = JsonDict("rowLabels")
   If initialCollection.Count > 0 Then
      For i = 1 To initialCollection.Count
          If TypeName(initialCollection(i)) = "Dictionary" Then
             Set initialDict = initialCollection(i)
             FieldName = initialDict("columnName")
             SortName = initialDict("sortType")
             fieldMap.Add FieldName, FieldName
             With PT.PivotFields(FieldName)
                .Orientation = xlRowField
                If UCase(SortName) = "DESC" Then
                   .AutoSort xlDescending, FieldName
                End If
             End With
          End If
      Next
   End If
End If
DoEvents
Progress.lblReport.Caption = "Row Labels Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

DoEvents
Progress.lblReport.Caption = "Adding Column Labels.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(56)
DoEvents
'Step-8: Pivot column labels.
If JsonDict.Exists("columnLabels") And LCase(TypeName(JsonDict("columnLabels"))) = "collection" Then
   Set initialCollection = JsonDict("columnLabels")
   If initialCollection.Count > 0 Then
      'initialCollection.Count To 1 Step -1 '''if we want labels in reverse order
      For i = 1 To initialCollection.Count
          If TypeName(initialCollection(i)) = "Dictionary" Then
             Set initialDict = initialCollection(i)
             FieldName = initialDict("columnName")
             fieldMap.Add FieldName, FieldName
             SortName = initialDict("sortType")
             With PT.PivotFields(FieldName)
                .Orientation = xlColumnField
                If UCase(SortName) = "DESC" Then
                   .AutoSort xlDescending, FieldName
                End If
             End With
          End If
      Next
   End If
End If
DoEvents
Progress.lblReport.Caption = "Column Labels Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

DoEvents
Progress.lblReport.Caption = "Adding Data Labels.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(63)
DoEvents
'Step-9: Pivot data labels.

If JsonDict.Exists("dataLabels") And LCase(TypeName(JsonDict("dataLabels"))) = "collection" Then
   Set initialCollection = JsonDict("dataLabels")
   If initialCollection.Count > 0 Then
      For i = 1 To initialCollection.Count
          If TypeName(initialCollection(i)) = "Dictionary" Then
             FormatStr = ""
             Set initialDict = initialCollection(i)
             FieldName = Trim(initialDict("columnName"))
             fieldMap.Add FieldName, FieldName & " "
             SortName = initialDict("sortType")
             numberFormat = initialDict("numberFormat")
             showNegative = initialDict("showNegativeAsRedColor")
             applyFormat = initialDict("applyNumberFormat")
             FormatStr = numberFormat
             If UCase(applyFormat) = "Y" Then
                If (Left(Trim(CStr(numberFormat)), 1) = "#" Or Mid(Trim(CStr(numberFormat)), 2, 1) = "#") And InStr(numberFormat, "%") = 0 Then
                   If UCase(showNegative) = "Y" Then
                      FormatStr = FormatStr & ";[Red](" & FormatStr & ")"
                   Else
                      FormatStr = FormatStr & ";-" & FormatStr
                   End If
                End If
                If InStr(numberFormat, "%") > 0 Then
                   If UCase(showNegative) = "Y" Then
                      FormatStr = FormatStr & ";[Red](" & FormatStr & ")"
                   Else
                      FormatStr = FormatStr & ";-" & FormatStr
                   End If
                   If InStr(numberFormat, "\") = 0 Then
                      FormatStr = Replace(FormatStr, "%", "\%")
                   Else
                      FormatStr = numberFormat
                   End If
                End If
             End If
             
             On Error Resume Next
             
             If initialDict("calculatedField") <> "" And Len(initialDict("calculatedField")) > 1 Then
                PT.CalculatedFields.Add FieldName, initialDict("calculatedField"), True
                With PT.PivotFields(FieldName)
                   .Orientation = xlDataField
                   .Function = XlTypeFnc(UCase(initialDict("aggregator")))
                   .Caption = FieldName & " "
                End With
             Else
                PT.AddDataField PT.PivotFields(FieldName), FieldName & " ", XlTypeFnc(UCase(initialDict("aggregator")))
             End If
             FieldName = FieldName & " "
             
             If initialDict("specialCalculation") <> "" And Len(initialDict("specialCalculation")) > 2 And LCase(initialDict("specialCalculation")) <> "none" Then
               If UCase(initialDict("specialCalculation")) = "PERCENTAGEOFGRANDTOTAL" Then
                  PT.PivotFields(FieldName).Calculation = xlPercentOfTotal
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOFCOLUMNTOTAL" Then
                  PT.PivotFields(FieldName).Calculation = xlPercentOfColumn
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOFROWTOTAL" Then
                  PT.PivotFields(FieldName).Calculation = xlPercentOfRow
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOFPARENTROWTOTAL" Then
                  PT.PivotFields(FieldName).Calculation = xlPercentOfParentRow
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOFPARENTROWCOUNT" Then
                  PT.PivotFields(FieldName).Calculation = xlPercentOfParentRow
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOFPARENTCOLUMNTOTAL" Then
                  PT.PivotFields(FieldName).Calculation = xlPercentOfParentColumn
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOFPARENTCOLUMNCOUNT" Then
                  PT.PivotFields(FieldName).Calculation = xlPercentOfParentColumn
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOF" Then
                  With PT.PivotFields(FieldName)
                      .Calculation = xlPercentOf
                      .BaseField = initialDict("baseField")
                      .BaseItem = initialDict("baseItem")
                  End With
               ElseIf UCase(initialDict("specialCalculation")) = "PERCENTAGEOFPARENTTOTAL" Then
                  With PT.PivotFields(FieldName)
                      .Calculation = xlPercentOfParent
                      .BaseField = initialDict("baseField")
                  End With
               End If
               If UCase(initialDict("specialCalculation")) <> "NONE" Then
                  If Len(FormatStr) > 0 Then
                     If InStr(FormatStr, "%") = 0 Then
                        If InStr(FormatStr, ";") > 0 Then
                           Dim formatvariant As Variant
                           formatvariant = Split(FormatStr, ";")
                           FormatStr = formatvariant(0) & " %" & ";" & formatvariant(1) & " %"
                        Else
                           FormatStr = FormatStr & " %"
                        End If
                     End If
                  Else
                     FormatStr = "0.00 %;-0.00 %"
                  End If
               End If
             End If
             
             If JsonDict.Exists("showZeroAsBlank") And UCase(JsonDict("showZeroAsBlank")) = "Y" Then
                FormatStr = FormatStr & ";;@"
             End If
             
             PT.PivotFields(FieldName).numberFormat = FormatStr
             
             Set DataRange = PT.PivotFields(FieldName).DataRange.SpecialCells(xlCellTypeVisible)
             
             If UCase(initialDict("hideColumn")) = "YES" Then
                DataRange.EntireColumn.Hidden = True
             Else
                DataRange.ColumnWidth = CDbl(initialDict("columnWidth"))
             End If
              On Error GoTo 0
          End If
      Next
   End If
End If
DoEvents
Progress.lblReport.Caption = "Data Labels Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

'Step-10: Adding Calculated Items.
DoEvents
Progress.lblReport.Caption = "Adding calculated items.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(68)
DoEvents
On Error Resume Next
If JsonDict.Exists("pivotCalsItems") And LCase(TypeName(JsonDict("pivotCalsItems"))) = "collection" Then
   Set initialCollection = JsonDict("pivotCalsItems")
   If initialCollection.Count > 0 Then
      For i = 1 To initialCollection.Count
          If TypeName(initialCollection(i)) = "Dictionary" Then
             Set initialDict = initialCollection(i)
             
             Dim itemColumnName As String
             Dim itemSpecialCalculation As String
             Dim itemBaseField As String
             
             FormatStr = ""
             itemColumnName = initialDict("columnName")
             itemSpecialCalculation = initialDict("specialCalculation")
             itemBaseField = initialDict("baseField")
             numberFormat = initialDict("numberFormat")
             FormatStr = numberFormat
             showNegative = initialDict("showNegativeAsRedColor")
             applyFormat = initialDict("applyNumberFormat")
             
             If UCase(applyFormat) = "Y" Then
                If (Left(Trim(CStr(numberFormat)), 1) = "#" Or Mid(Trim(CStr(numberFormat)), 2, 1) = "#") And InStr(numberFormat, "%") = 0 Then
                   If UCase(showNegative) = "Y" Then
                      FormatStr = FormatStr & ";[Red](" & FormatStr & ")"
                   Else
                      FormatStr = FormatStr & ";-" & FormatStr
                   End If
                End If
                If InStr(numberFormat, "%") > 0 Then
                   If UCase(showNegative) = "Y" Then
                      FormatStr = FormatStr & ";[Red](" & FormatStr & ")"
                   Else
                      FormatStr = FormatStr & ";-" & FormatStr
                   End If
                   If InStr(numberFormat, "\") = 0 Then
                      FormatStr = Replace(FormatStr, "%", "\%")
                   Else
                      FormatStr = numberFormat
                   End If
                End If
             End If
             
             PT.PivotFields(itemBaseField).CalculatedItems.Add itemColumnName, itemSpecialCalculation, True
             If FormatStr = "" Then
                FormatStr = "General"
             End If
             PT.PivotFields(itemColumnName).numberFormat = numberFormat
          End If
      Next
   End If
End If
On Error GoTo 0

'Step-11: Grand Totals and Sub-totals.
DoEvents
Progress.lblReport.Caption = "Applying Grand Totals and Sub-Totals.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(70)
DoEvents
Call Pivot_Subtotals(JsonDict, PT)
DoEvents
Progress.lblReport.Caption = "Completed Totals.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

DoEvents
Progress.lblReport.Caption = "Applying Number Formats.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(77)
DoEvents
'Step-12: Row labels and Column labels number format.
Call FormatPivotRowsAndColumns(JsonDict, PT)
DoEvents
Progress.lblReport.Caption = "Formatting Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

' --- Moved here (was after Step-14, near the end of this sub) ---
' Row groups are expanded/collapsed to their FINAL shape before conditional formatting and
' filters are applied below, so Excel never has to reconcile already-registered
' FormatConditions/PivotFilters against a row range that changes size/shape afterward. That
' reconciliation (when expand ran last, as originally written) is what was making
' ExpandCollapseDetails grind at ~30% CPU for a long time on reports with a high-cardinality
' row field (e.g. an Order ID) - moving this earlier means there is nothing yet registered
' that needs to be migrated when the expand happens.
DoEvents
Progress.lblReport.Caption = "Applying Auto Expand/Collapse.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call DebugLog.LogMsg("MainModule.Generating_Pivots", "BEFORE autoExpandALL block")
If JsonDict.Exists("autoExpandALL") And LCase(TypeName(JsonDict("autoExpandALL"))) = "string" Then
   If LCase(JsonDict("autoExpandALL")) = "true" Then
      If JsonDict.Exists("noOfRowGroupsExpand") Then
         If CStr(JsonDict("noOfRowGroupsExpand")) = "-1" Then
            Call DebugLog.LogMsg("MainModule.Generating_Pivots", "BEFORE ExpandCollapseDetails(True)")
            Call ExpandCollapseDetails(PT, True)
            Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER ExpandCollapseDetails(True)")
         ElseIf CStr(JsonDict("noOfRowGroupsExpand")) = "0" Then
            Call DebugLog.LogMsg("MainModule.Generating_Pivots", "BEFORE ExpandCollapseDetails(False)")
            Call ExpandCollapseDetails(PT, False)
            Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER ExpandCollapseDetails(False)")
         Else
            Call DebugLog.LogMsg("MainModule.Generating_Pivots", "BEFORE ExpandCollapseFieldDetails(" & JsonDict("noOfRowGroupsExpand") & ")")
            Call ExpandCollapseFieldDetails(PT, CInt(JsonDict("noOfRowGroupsExpand")))
            Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER ExpandCollapseFieldDetails")
         End If
      End If
   End If
End If
Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER autoExpandALL block")
DoEvents

DoEvents
Progress.lblReport.Caption = "Applying Conditional Formats.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
Call ProgressBar_Chart(84)
DoEvents
'Step-13: Applying conditional formats.

If JsonDict.Exists("pivotConditionalFormatting") And LCase(TypeName(JsonDict("pivotConditionalFormatting"))) = "string" And _
   UCase(JsonDict("pivotConditionalFormatting")) = "Y" Then
   Dim i1 As Integer
   Dim CFinitialCollection  As Collection
   Dim CFinitialDict As Dictionary
   
   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "Step-13: pivotConditionalFormatting=Y, starting rowLabels CF pass")
   
   If JsonDict.Exists("rowLabels") And LCase(TypeName(JsonDict("rowLabels"))) = "collection" Then
      Set initialCollection = JsonDict("rowLabels")
      If initialCollection.Count > 0 Then
         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "rowLabels count=" & initialCollection.Count)
         For i = 1 To initialCollection.Count
             If LCase(TypeName(initialCollection(i))) = "dictionary" Then
                Set initialDict = initialCollection(i)
                FieldName = initialDict("columnName")
                Call DebugLog.LogMsg("MainModule.Generating_Pivots", "rowLabels[" & i & "] field=[" & FieldName & "] - checking conditionalFormat")
                If initialDict.Exists("conditionalFormat") Then
                   If LCase(TypeName(initialDict("conditionalFormat"))) = "collection" Then
                      Set CFinitialCollection = initialDict("conditionalFormat")
                      If CFinitialCollection.Count > 0 Then
                         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "rowLabels[" & i & "] field=[" & FieldName & "] conditionalFormat rule count=" & CFinitialCollection.Count)
                         For i1 = 1 To CFinitialCollection.Count
                             If LCase(TypeName(CFinitialCollection(i1))) = "dictionary" Then
                                Set CFinitialDict = CFinitialCollection(i1)
                                If CFinitialDict.Count > 0 Then
                                   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "rowLabels[" & i & "] field=[" & FieldName & "] rule[" & i1 & "] BEFORE CF.PivotCF")
                                   Call CF.PivotCF(PT, CFinitialDict, FieldName, True)
                                   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "rowLabels[" & i & "] field=[" & FieldName & "] rule[" & i1 & "] AFTER CF.PivotCF")
                                End If
                             End If
                         Next
                      End If
                   ElseIf LCase(TypeName(initialDict("conditionalFormat"))) = "dictionary" Then
                      Set initialDict1 = initialDict("conditionalFormat")
                      If initialDict1.Count > 0 Then
                         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "rowLabels[" & i & "] field=[" & FieldName & "] single-rule BEFORE CF.PivotCF")
                         Call CF.PivotCF(PT, initialDict1, FieldName, True)
                         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "rowLabels[" & i & "] field=[" & FieldName & "] single-rule AFTER CF.PivotCF")
                      End If
                   End If
                End If
             End If
         Next
      End If
   End If
   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "Step-13: rowLabels CF pass complete, starting dataLabels CF pass")
   If JsonDict.Exists("dataLabels") And LCase(TypeName(JsonDict("dataLabels"))) = "collection" Then
      Set initialCollection = JsonDict("dataLabels")
      If initialCollection.Count > 0 Then
         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "dataLabels count=" & initialCollection.Count)
         For i = 1 To initialCollection.Count
             If LCase(TypeName(initialCollection(i))) = "dictionary" Then
                Set initialDict = initialCollection(i)
                FieldName = initialDict("columnName")
                Call DebugLog.LogMsg("MainModule.Generating_Pivots", "dataLabels[" & i & "] field=[" & FieldName & "] - checking conditionalFormat")
                If initialDict.Exists("conditionalFormat") Then
                   If LCase(TypeName(initialDict("conditionalFormat"))) = "collection" Then
                      Set CFinitialCollection = initialDict("conditionalFormat")
                      If CFinitialCollection.Count > 0 Then
                         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "dataLabels[" & i & "] field=[" & FieldName & "] conditionalFormat rule count=" & CFinitialCollection.Count)
                         For i1 = 1 To CFinitialCollection.Count
                             If LCase(TypeName(CFinitialCollection(i1))) = "dictionary" Then
                                Set CFinitialDict = CFinitialCollection(i1)
                                If CFinitialDict.Count > 0 Then
                                   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "dataLabels[" & i & "] field=[" & FieldName & "] rule[" & i1 & "] BEFORE CF.PivotCF")
                                   Call CF.PivotCF(PT, CFinitialDict, FieldName & " ", False)
                                   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "dataLabels[" & i & "] field=[" & FieldName & "] rule[" & i1 & "] AFTER CF.PivotCF")
                                End If
                             End If
                         Next
                      End If
                   ElseIf LCase(TypeName(initialDict("conditionalFormat"))) = "dictionary" Then
                      Set initialDict1 = initialDict("conditionalFormat")
                      If initialDict1.Count > 0 Then
                         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "dataLabels[" & i & "] field=[" & FieldName & "] single-rule BEFORE CF.PivotCF")
                         Call CF.PivotCF(PT, initialDict1, FieldName & " ", False)
                         Call DebugLog.LogMsg("MainModule.Generating_Pivots", "dataLabels[" & i & "] field=[" & FieldName & "] single-rule AFTER CF.PivotCF")
                      End If
                   End If
                End If
             End If
         Next
      End If
   End If
   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "Step-13: dataLabels CF pass complete")
End If


DoEvents
Progress.lblReport.Caption = "Conditional Formats Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

'Step-14: Applying pivot value and label filters.
DoEvents
Progress.lblReport.Caption = "Applying pivot value and label filters.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents
Call PivotRowsAndColumnsFilters(JsonDict, PT)
DoEvents
Progress.lblReport.Caption = "Applying pivot value and label filters Completed.."
Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
DoEvents

If JsonDict.Exists("nullDisplay") And LCase(TypeName(JsonDict("nullDisplay"))) = "string" And _
   JsonDict("nullDisplay") <> "" Then
   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "BEFORE PT.NullString assignment")
   PT.NullString = JsonDict("nullDisplay")
   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER PT.NullString assignment")
End If

Call DebugLog.LogMsg("MainModule.Generating_Pivots", "BEFORE pivotProtection block")
If JsonDict.Exists("pivotProtection") And LCase(TypeName(JsonDict("pivotProtection"))) = "string" Then
   ' EnableFieldList/EnableFieldDialog attach/show the PivotTable Field List task pane, which
   ' needs an actual screen paint. The whole macro runs with ScreenUpdating=False (see
   ' Speedup_Excel), so Excel can end up waiting on a repaint that's been suppressed - a
   ' deterministic deadlock, confirmed reproducible on a completely fresh file at this exact
   ' call. Restore ScreenUpdating just for this one call, then put it back.
   Application.ScreenUpdating = True
   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "ScreenUpdating temporarily restored for pivotProtection block")
   If LCase(JsonDict("pivotProtection")) = "y" Then
      DoEvents
      Progress.lblReport.Caption = "Protecting Pivot Table and Layout.."
      Call DebugLog.LogMsg("MainModule.Generating_Pivots", Progress.lblReport.Caption)
      DoEvents
      Call DebugLog.LogMsg("MainModule.Generating_Pivots", "BEFORE RestrictPivotTable")
      Call RestrictPivotTable(PT)
      Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER RestrictPivotTable, BEFORE Protecting_Sheet")
      Call Protecting_Sheet(PivotSheet)
      Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER Protecting_Sheet")
   Else
      ' AllowPivotTable(PT) deliberately NOT called here. PT at this point is a PivotTable
      ' freshly created via PivotSheet.PivotTables.Add just above in this same sub - it already
      ' defaults to EnableDrilldown/EnableFieldList/EnableFieldDialog/PivotCache.EnableRefresh
      ' all True and every DragTo* True, since nothing in this run has restricted it (that only
      ' happens in the "y" branch above, via RestrictPivotTable). Calling AllowPivotTable here
      ' was reasserting settings the pivot already has by construction, and is the exact call
      ' that has hung in every round of testing since it was reached (Rounds 4-9) - skipping it
      ' removes the call entirely rather than continuing to patch around whichever property
      ' inside it blocks next.
      Call DebugLog.LogMsg("MainModule.Generating_Pivots", "SKIPPED AllowPivotTable (redundant on a freshly-created, already-unrestricted PivotTable)")
   End If
   Application.ScreenUpdating = False
   Call DebugLog.LogMsg("MainModule.Generating_Pivots", "ScreenUpdating turned back off after pivotProtection block")
End If
Call DebugLog.LogMsg("MainModule.Generating_Pivots", "AFTER pivotProtection block")

Set DataRange = Nothing
Set DummyRng = Nothing
Set initialCollection = Nothing
Set initialDict = Nothing
Set initialDict1 = Nothing
Set PCache = Nothing
Set PT = Nothing
On Error GoTo 0

Call DebugLog.LogMsg("MainModule.Generating_Pivots", "=== Generating_Pivots end ===")

End Sub
Private Sub ExpandCollapseFieldDetails(ByVal OrbPivot As PivotTable, ByVal loopint As Integer)
    On Error Resume Next
    Dim pf As PivotField
    Dim i As Integer
    Dim iFieldCount As Long
    
    Call DebugLog.LogMsg("MainModule.ExpandCollapseFieldDetails", "ENTER loopint=" & loopint & " RowFields.Count=" & OrbPivot.RowFields.Count)
    Call ExpandCollapseDetails(OrbPivot, False)
    
    iFieldCount = OrbPivot.RowFields.Count - 1
    
    For i = iFieldCount To 1 Step -1
       For Each pf In OrbPivot.RowFields
          If pf.Position <= loopint Then
             If pf.ShowDetail = False Then
                Call CF.SetObjPropertyWithRetry(pf, "ShowDetail", True, "ExpandCollapseFieldDetails field=[" & pf.Name & "]")
             End If
          End If
       Next
    Next
    Call DebugLog.LogMsg("MainModule.ExpandCollapseFieldDetails", "EXIT")
    
    On Error GoTo 0
End Sub
Private Sub ExpandCollapseDetails(ByVal OrbPivot As PivotTable, ByVal ShowHide As Boolean)
    On Error Resume Next
    Dim pf As PivotField
    Call DebugLog.LogMsg("MainModule.ExpandCollapseDetails", "ENTER ShowHide=" & ShowHide & " RowFields.Count=" & OrbPivot.RowFields.Count)
    For Each pf In OrbPivot.RowFields
       Call CF.SetObjPropertyWithRetry(pf, "ShowDetail", ShowHide, "ExpandCollapseDetails field=[" & pf.Name & "]")
    Next
    Call DebugLog.LogMsg("MainModule.ExpandCollapseDetails", "EXIT")
    On Error GoTo 0
End Sub
Private Sub FormatPivotRowsAndColumns(ByVal JsonDict As Dictionary, ByVal PT As PivotTable)

Dim initialCollection As Collection

On Error Resume Next

If JsonDict.Exists("rowLabels") And LCase(TypeName(JsonDict("rowLabels"))) = "collection" Then
   Set initialCollection = JsonDict("rowLabels")
   If initialCollection.Count > 0 Then
      Call FormatPivotRowsAndColumns_Extend(initialCollection, PT, JsonDict)
   End If
End If

If JsonDict.Exists("columnLabels") And LCase(TypeName(JsonDict("columnLabels"))) = "collection" Then
   Set initialCollection = JsonDict("columnLabels")
   If initialCollection.Count > 0 Then
       Call FormatPivotRowsAndColumns_Extend(initialCollection, PT, JsonDict)
   End If
End If
Set initialCollection = Nothing
On Error GoTo 0
End Sub
Private Sub PivotRowsAndColumnsFilters(ByVal JsonDict As Dictionary, ByVal PT As PivotTable)

Dim initialCollection As Collection

On Error Resume Next

If JsonDict.Exists("columnLabels") And LCase(TypeName(JsonDict("columnLabels"))) = "collection" Then
   Set initialCollection = JsonDict("columnLabels")
   If initialCollection.Count > 0 Then
       Call PivotRowsAndColumnsFilters_Extend(initialCollection, PT, JsonDict)
   End If
End If

If JsonDict.Exists("rowLabels") And LCase(TypeName(JsonDict("rowLabels"))) = "collection" Then
   Set initialCollection = JsonDict("rowLabels")
   If initialCollection.Count > 0 Then
      Call PivotRowsAndColumnsFilters_Extend(initialCollection, PT, JsonDict)
   End If
End If

Set initialCollection = Nothing
On Error GoTo 0
End Sub
Private Sub PivotRowsAndColumnsFilters_Extend(ByVal initialCollection As Collection, ByVal PT As PivotTable, ByVal JsonDict As Dictionary)

Dim initialDict As Dictionary
Dim initialDict1 As Dictionary
Dim FormatStr As String
Dim i As Integer
Dim FieldName As String

On Error Resume Next

For i = 1 To initialCollection.Count
    If TypeName(initialCollection(i)) = "Dictionary" Then
       Set initialDict = initialCollection(i)
       FieldName = initialDict("columnName")
       If JsonDict.Exists("pivotLabelFilter") And LCase(TypeName(JsonDict("pivotLabelFilter"))) = "string" And UCase(JsonDict("pivotLabelFilter")) = "Y" Then
          If initialDict.Exists("labelFilters") And LCase(TypeName(initialDict("labelFilters"))) = "dictionary" Then
             Set initialDict1 = initialDict("labelFilters")
             If initialDict1.Count > 0 Then
                Call PivotFilters.PivotFilters(PT, initialDict1, FieldName, "label")
              End If
          End If
       End If
       If JsonDict.Exists("pivotValueFilter") And LCase(TypeName(JsonDict("pivotValueFilter"))) = "string" And UCase(JsonDict("pivotValueFilter")) = "Y" Then
          If initialDict.Exists("valueFilters") And LCase(TypeName(initialDict("valueFilters"))) = "dictionary" Then
             Set initialDict1 = initialDict("valueFilters")
             If initialDict1.Count > 0 Then
                Call PivotFilters.PivotFilters(PT, initialDict1, FieldName, "value")
             End If
          End If
       End If
    End If
Next

Set initialDict = Nothing
On Error GoTo 0
End Sub
Private Sub FormatPivotRowsAndColumns_Extend(ByVal initialCollection As Collection, ByVal PT As PivotTable, ByVal JsonDict As Dictionary)

Dim initialDict As Dictionary
Dim FormatStr As String
Dim i As Integer
Dim FieldName As String

On Error Resume Next

For i = 1 To initialCollection.Count
    If TypeName(initialCollection(i)) = "Dictionary" Then
       Set initialDict = initialCollection(i)
       If UCase(initialDict("applyNumberFormat")) = "Y" Then
          FieldName = initialDict("columnName")
          FormatStr = ""
          FormatStr = initialDict("numberFormat")
          FormatStr = Replace(FormatStr, "EEEE", "dddd", , , vbTextCompare)
       
          With PT.PivotFields(FieldName)
             If InStr(FormatStr, "AM/PM") > 0 Then
                .numberFormat = FormatStr
             Else
                If InStr(FormatStr, "/") > 0 Then
                   .numberFormat = Replace(FormatStr, "/", "\/") & ";@"
                Else
                   .numberFormat = FormatStr
                End If
             End If
          End With
          If JsonDict.Exists("pivotSubTotals") And LCase(TypeName(JsonDict("pivotSubTotals"))) = "string" Then 'This is sub-total section
             If UCase(JsonDict("pivotSubTotals")) <> "HIDE" Then
                If initialDict.Exists("showTotal") Then
                   If UCase(initialDict("showTotal")) = "HIDE" Then
                      Call HideSubTotals(PT, FieldName)
                   End If
                End If
             End If
          End If
       End If
    End If
Next

Set initialDict = Nothing
On Error GoTo 0
End Sub
Private Sub Format_RawSheet(ByVal sht As Worksheet, ByVal rg As Range, ByVal ColName As String, _
                            ByVal FormatStr As String, ByVal dataType As String)

Dim MatchCol As Integer
Dim RowsLong As Long
Dim i As Integer
Dim rng As Range

On Error Resume Next

sht.Activate
RowsLong = rg.Rows.Count

MatchCol = Application.WorksheetFunction.Match(ColName, sht.Range("A1:ZZ1"), 0)
If MatchCol > 0 Then
   Set rng = sht.Range(sht.Cells(2, MatchCol), sht.Cells(RowsLong, MatchCol))
   'Rng = Evaluate(Rng.Address & "*1")
   rng.Name = "orbformat"
   rng = Evaluate("IF(LEN(orbformat)=0,"""",orbformat*1)")
   
   If InStr(FormatStr, "AM/PM") > 0 Then
      rng.numberFormat = FormatStr
   Else
      If InStr(FormatStr, "/") > 0 Then
         rng.numberFormat = Replace(FormatStr, "/", "\/") & ";@"
      Else
         rng.numberFormat = FormatStr
      End If
   End If
   
'   If UCase(dataType) = "DATE" Then
'      rng.numberFormat = "dd\/mm\/yyyy;@"
'   ElseIf UCase(dataType) = "DATETIME" Then
'      rng.numberFormat = "m/d/yyyy h:mm"
'   Else
'      rng.numberFormat = FormatStr
'   End If
   
   ActiveWorkbook.names("orbformat").Delete
End If
Set rng = Nothing
On Error GoTo 0
End Sub
Private Function XlTypeFnc(ByVal fType As String) As Long
Select Case UCase(fType)
  Case "SUM"
        XlTypeFnc = -4157
  Case "COUNT", "COUNT_DISTINCT"
        XlTypeFnc = -4112
  Case "AVG", "AVERAGE"
        XlTypeFnc = -4106
  Case "MAX"
        XlTypeFnc = -4136
  Case "MIN"
        XlTypeFnc = -4139
  Case "PRODUCT"
        XlTypeFnc = -4149
  Case "COUNT NUMBERS"
        XlTypeFnc = -4113
  Case "STDDEV"
        XlTypeFnc = -4155
  Case "STDDEVP"
        XlTypeFnc = -4156
  Case "VAR", "VARIANCE"
        XlTypeFnc = -4164
  Case "VARP", "VARIANCEP"
        XlTypeFnc = -4165
  Case Else
        XlTypeFnc = -4157
End Select
End Function
Private Sub Pivot_Subtotals(ByVal JsonDict As Variant, ByVal OrbitPivot As PivotTable)
        
On Error Resume Next

If JsonDict.Exists("pivotSubTotals") And LCase(TypeName(JsonDict("pivotSubTotals"))) = "string" Then 'This is sub-total section
    If UCase(JsonDict("pivotSubTotals")) = "FIRST" Then
        Call TurnOnTotals(OrbitPivot)
        OrbitPivot.SubtotalLocation xlAtTop
    ElseIf UCase(JsonDict("pivotSubTotals")) = "HIDE" Then
        Call TurnOffTotals(OrbitPivot)
    ElseIf UCase(JsonDict("pivotSubTotals")) = "LAST" Then
        Call TurnOnTotals(OrbitPivot)
        OrbitPivot.SubtotalLocation xlAtBottom
    End If
End If

If JsonDict.Exists("pivotGrandTotals") And LCase(TypeName(JsonDict("pivotGrandTotals"))) = "string" Then 'This is Grand Totals Section
    If UCase(JsonDict("pivotGrandTotals")) = "FIRST" Or UCase(JsonDict("pivotGrandTotals")) = "LAST" Then
        With ActiveSheet.PivotTables("OrbitPivot")
             .ColumnGrand = True
             .RowGrand = True
        End With
    ElseIf UCase(JsonDict("pivotGrandTotals")) = "HIDE" Then
        With ActiveSheet.PivotTables("OrbitPivot")
             .ColumnGrand = False
             .RowGrand = False
        End With
    End If
End If

If JsonDict.Exists("pivotRowGrandTotals") And LCase(TypeName(JsonDict("pivotRowGrandTotals"))) = "string" Then 'This is row grand totals
    If UCase(JsonDict("pivotRowGrandTotals")) = "FIRST" Then
        OrbitPivot.RowGrand = True
    ElseIf UCase(JsonDict("pivotRowGrandTotals")) = "HIDE" Then
        OrbitPivot.RowGrand = False
    ElseIf UCase(JsonDict("pivotRowGrandTotals")) = "LAST" Then
        OrbitPivot.RowGrand = True
    End If
End If
If JsonDict.Exists("pivotColumnGrandTotals") And LCase(TypeName(JsonDict("pivotColumnGrandTotals"))) = "string" Then 'This is column grand totals
    If UCase(JsonDict("pivotColumnGrandTotals")) = "FIRST" Then
        OrbitPivot.ColumnGrand = True
    ElseIf UCase(JsonDict("pivotColumnGrandTotals")) = "HIDE" Then
        OrbitPivot.ColumnGrand = False
    ElseIf UCase(JsonDict("pivotColumnGrandTotals")) = "LAST" Then
        OrbitPivot.ColumnGrand = True
    End If
End If

On Error GoTo 0
End Sub
Private Sub TurnOffTotals(ByVal OrbPivot As PivotTable)
    On Error Resume Next
    Dim pf As PivotField
    For Each pf In OrbPivot.PivotFields
        pf.Subtotals = Array( _
        False, False, False, False, False, False, False, False, False, False, False, False)
        'pf.Subtotals(1) = False
    Next
    On Error GoTo 0
End Sub
Private Sub TurnOnTotals(ByVal OrbPivot As PivotTable)
    On Error Resume Next
    Dim pf As PivotField
    For Each pf In OrbPivot.PivotFields
        pf.Subtotals = Array( _
        True, False, False, False, False, False, False, False, False, False, False, False)
        'pf.Subtotals(1) = True
    Next
    On Error GoTo 0
End Sub
Private Sub HideSubTotals(ByVal OrbPivot As PivotTable, ByVal Fld As String)
    On Error Resume Next
    Dim pf As PivotField
    For Each pf In OrbPivot.PivotFields
        If UCase(pf.Name) = UCase(Trim(Fld)) Then
           pf.Subtotals = Array( _
           False, False, False, False, False, False, False, False, False, False, False, False)
           'pf.Subtotals(1) = False
           Exit For
        End If
    Next
    On Error GoTo 0
End Sub
Private Sub Protecting_Sheet(ByVal sht As Worksheet)
On Error Resume Next
Call DebugLog.LogMsg("MainModule.Protecting_Sheet", "ENTER sheet=[" & sht.Name & "]")
sht.Protect Password:="Orb1t!@#", DrawingObjects:=False, Contents:=True, Scenarios:= _
        False, AllowFormattingColumns:=True, AllowFormattingRows:=True, _
        AllowUsingPivotTables:=True
Call DebugLog.LogMsg("MainModule.Protecting_Sheet", "EXIT (Err=" & Err.Number & ")")
On Error GoTo 0
End Sub
Sub UnProtecting_Sheet(ByVal sht As Worksheet)
On Error Resume Next
Call DebugLog.LogMsg("MainModule.UnProtecting_Sheet", "ENTER sheet=[" & sht.Name & "]")
sht.Unprotect Password:="Orb1t!@#"
Call DebugLog.LogMsg("MainModule.UnProtecting_Sheet", "EXIT (Err=" & Err.Number & ")")
On Error GoTo 0
End Sub
Private Function CheckIfSheetIsProtected(ByVal sht As Worksheet) As Boolean

If sht.ProtectContents = True Then
   CheckIfSheetIsProtected = True
Else
   CheckIfSheetIsProtected = False
End If
End Function
Private Sub RestrictPivotTable(ByVal PT As PivotTable)
Dim pf As PivotField
Dim PC As PivotCache

On Error Resume Next

Call DebugLog.LogMsg("MainModule.RestrictPivotTable", "ENTER PivotFields.Count=" & PT.PivotFields.Count)

Call CF.SetObjPropertyWithRetry(PT, "EnableDrilldown", False, "RestrictPivotTable")
' EnableFieldList/EnableFieldDialog deliberately NOT set here - confirmed (across multiple
' fresh-file, no-add-in, ScreenUpdating-on/off reproductions) to block indefinitely on this
' machine. They only control two authoring-time UI conveniences (the Field List pane, the
' Field Settings dialog), not report correctness or end-user interaction, so they're skipped
' rather than chasing the exact COM-level cause further.

' PT.PivotCache is a property GET (returns the PivotCache object) - split from the EnableRefresh
' SET so the log can tell which one actually blocks, since a GET blocking would never even reach
' SetObjPropertyWithRetry (the block happens while evaluating the argument, before the call).
Call DebugLog.LogMsg("MainModule.RestrictPivotTable", "PivotCaches.Count=" & ThisWorkbook.PivotCaches.Count & " BEFORE PT.PivotCache GET")
Set PC = PT.PivotCache
Call DebugLog.LogMsg("MainModule.RestrictPivotTable", "AFTER PT.PivotCache GET")
' Yield here too - the GET doesn't go through SetObjPropertyWithRetry, so without this the very
' next call (EnableRefresh) fires with no gap after this one succeeds, same pattern as before.
DoEvents
Sleep 50
Call CF.SetObjPropertyWithRetry(PC, "EnableRefresh", False, "RestrictPivotTable PivotCache")

For Each pf In PT.PivotFields
    Call DebugLog.LogMsg("MainModule.RestrictPivotTable", "field=[" & pf.Name & "] BEFORE DragTo* writes")
    Call CF.SetObjPropertyWithRetry(pf, "DragToPage", False, "RestrictPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToRow", False, "RestrictPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToColumn", False, "RestrictPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToData", False, "RestrictPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToHide", False, "RestrictPivotTable field=[" & pf.Name & "]")
Next pf

Call DebugLog.LogMsg("MainModule.RestrictPivotTable", "EXIT")

On Error GoTo 0

End Sub
Private Sub AllowPivotTable(ByVal PT As PivotTable)

Dim pf As PivotField
Dim PC As PivotCache
On Error Resume Next

Call DebugLog.LogMsg("MainModule.AllowPivotTable", "ENTER PivotFields.Count=" & PT.PivotFields.Count)

Call CF.SetObjPropertyWithRetry(PT, "EnableDrilldown", True, "AllowPivotTable")
' EnableFieldList/EnableFieldDialog deliberately NOT set here - see matching note in
' RestrictPivotTable above. Confirmed to block indefinitely; skipped as a pragmatic fix.

' PT.PivotCache is a property GET - split from the EnableRefresh SET, see matching comment in
' RestrictPivotTable above.
Call DebugLog.LogMsg("MainModule.AllowPivotTable", "PivotCaches.Count=" & ThisWorkbook.PivotCaches.Count & " BEFORE PT.PivotCache GET")
Set PC = PT.PivotCache
Call DebugLog.LogMsg("MainModule.AllowPivotTable", "AFTER PT.PivotCache GET")
' Yield here too - see matching comment in RestrictPivotTable above.
DoEvents
Sleep 50
Call CF.SetObjPropertyWithRetry(PC, "EnableRefresh", True, "AllowPivotTable PivotCache")

For Each pf In PT.PivotFields
    Call DebugLog.LogMsg("MainModule.AllowPivotTable", "field=[" & pf.Name & "] BEFORE DragTo* writes")
    Call CF.SetObjPropertyWithRetry(pf, "DragToPage", True, "AllowPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToRow", True, "AllowPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToColumn", True, "AllowPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToData", True, "AllowPivotTable field=[" & pf.Name & "]")
    Call CF.SetObjPropertyWithRetry(pf, "DragToHide", True, "AllowPivotTable field=[" & pf.Name & "]")
Next pf

Call DebugLog.LogMsg("MainModule.AllowPivotTable", "EXIT")

On Error GoTo 0
End Sub
```

---

## Once you have a log

Paste (or attach) the tail of `PublisherReport_Debug.log` from the hung run back to me — the
last ~20-30 lines are usually enough — and I'll turn that into a targeted fix instead of another
round of theories. If the hang reproduces reliably on a specific report, it'd also help to know
roughly how many rows are in `RawData` and how many distinct row-field/column-field combinations
the report has (i.e. how "wide" the pivot's row/column hierarchy is), since that governs how big
`pf.DataRange.Areas.Count` can get.

## Round 10 — past every prior blocker; new crash in previously-untouched territory

`Generating_Pivots` completed **entirely** this run (`=== Generating_Pivots end ===` logged) -
the Round 9 fix worked. The run then crashed one step later, in `Start_Pivot_Generation`'s final
`"Other Miscellaneous.."` section (Step-14: the optional `saveXlsmToXlsx`/`deleteXlsmFile`
`ActiveWorkbook.SaveAs`/file-delete logic), which had never been instrumented before since every
prior round never got this far.

Added logging around: the `saveXlsmToXlsx` check itself, immediately before/after
`ActiveWorkbook.SaveAs`, immediately before/after `FSO.Deletefile`, the final
`ConfigSheet.Range("A3")` write, `Application.GoTo`, and the closing `Speedup_Excel(False)`. Since
this was an actual **crash** (not a hang) unlike every previous round, it's also worth checking
Windows' own crash record if Excel offered a "Send feedback"/error-reporting prompt, or checking
**Event Viewer → Windows Logs → Application** for an `.appcrash` entry for `EXCEL.EXE` around the
crash time - that record's fault module (e.g. a specific DLL name) could independently confirm
or contradict what the new log narrows down to.

**Please test this version and send the new log** (plus the Event Viewer crash entry if one
exists) — this is genuinely new ground, not a repeat of anything seen so far.
