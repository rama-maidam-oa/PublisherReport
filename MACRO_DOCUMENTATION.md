# VBA Macro Documentation — `default_macro_pivot_template.xlsm`

This workbook is an **Orbit BI / Orbit Analytics** template. Its VBA project builds a fully
formatted, formula-driven Excel **PivotTable report** on the `ReportOutput` sheet, entirely
controlled by a **JSON configuration string** stored in a hidden `PivotOrbConfig` sheet, using
raw tabular data found on a `RawData` sheet. No user interaction is required — the whole report
is generated automatically the first time the `ReportOutput` sheet is opened/activated.

## Required workbook sheets

| Sheet name        | Purpose                                                                                                   |
|--------------------|-------------------------------------------------------------------------------------------------------------|
| `RawData`          | Raw source data (table with headers in row 1) used to build the PivotCache/PivotTable.                    |
| `PivotOrbConfig`   | Hidden config sheet. `A1` = JSON string describing the report layout. `A3` = execution flag (`"Y"` once the macro has already run, prevents re-running). |
| `ReportOutput`     | Destination sheet where the PivotTable (`OrbitPivot`) is built.                                            |

## Entry point / trigger

**`Sheet1.Worksheet_Activate`** (code-behind of the sheet whose `.Name = "ReportOutput"`):

```vba
Private Sub Worksheet_Activate()
    If sht.Name = "ReportOutput" And UCase(PivotOrbConfig!A3) <> "Y" Then
        Call MainModule.Start_Pivot_Generation
    End If
End Sub
```

Every time the `ReportOutput` sheet becomes active, this checks whether the report has already
been generated (`PivotOrbConfig!A3 = "Y"`). If not, it kicks off the whole pivot-generation
pipeline. Once generation succeeds, `MainModule` sets `PivotOrbConfig!A3 = "Y"` so the macro
never runs again for this file (re-activating the sheet afterwards does nothing).

`Sheet2`, `Sheet3`, `Sheet4`, and `ThisWorkbook` contain no code — only default class attributes.

## VBA Project modules

| Module | Type | Role |
|---|---|---|
| `Sheet1` | Sheet class | Fires the entry point (`Worksheet_Activate`) on `ReportOutput`. |
| `MainModule` | Standard module | Orchestrates the entire pivot build/validation/format pipeline. |
| `CF` | Standard module | Builds Excel conditional-formatting rules on pivot fields from JSON rule specs. |
| `PivotFilters` | Standard module | Applies pivot **label filters** and **value filters** from JSON rule specs. |
| `JsonConverter` | Standard module | Third-party **VBA-JSON** library (Tim Hall, MIT license) + **VBA-UTC** date helpers. Parses/serializes JSON. |
| `Progress` | UserForm | Modeless progress-bar dialog (`Bar`, `Border`, `Text`, `lblReport` controls) shown while the report builds. Has no code-behind — driven entirely from `MainModule`. |

---

## `MainModule` — the orchestrator

### `Start_Pivot_Generation` (Public Sub — the main routine)

Runs top‑to‑bottom, updating the `Progress` form's caption/percentage bar at each step
(`ProgressBar_Chart`). Wrapped by `Speedup_Excel(True/False)`, which disables screen updating,
alerts, automatic calculation and events during the run (turned back on at the end) for
performance.

Step by step:

1. **Validate the workbook file name.** Extracts the file name from `ThisWorkbook.FullName` and
   checks it with `ValidFileName` against the character set `\ / : * ? < > | [ ] "`. If invalid,
   shows a message box naming the offending character and aborts.
2. **Validate `RawData` sheet.** Must exist and contain data (last used row/column found via
   `Cells.Find`). Aborts (with message) if the sheet is missing, has no data, or its `A1` cell
   still contains the unrendered template placeholder `"${TableHeader}"`.
3. **Validate `PivotOrbConfig` sheet.** Must exist and `A1` must be non-empty (this holds the
   JSON config).
4. **Validate `ReportOutput` sheet.** Must exist (this is where the pivot gets built).
5. **Check the "already executed" flag.** If `PivotOrbConfig!A3 = "Y"`, shows a message box
   telling the user to clear it manually, then aborts (idempotency guard).
6. **Parse JSON config.** `JsonConverter.ParseJSON(PivotOrbConfig!A1)` into a `Dictionary`. If
   parsing fails / returns Nothing, shows "Invalid JSON..." and aborts.
7. **Format the raw data sheet** (`RawSheetFormat`) — if the JSON has a `"rawData"` key, applies
   per-column number formats to `RawData` (see below).
8. **Build the pivot table** — calls `Generating_Pivots` (see below), which does the bulk of the
   report construction (layout, fields, calculated fields/items, totals, number formats,
   conditional formatting, filters, expand/collapse, null display, protection).
9. **Optional save-as-xlsx.** If JSON `"saveXlsmToXlsx" = "Y"`: saves the active workbook as
   `.xlsx` (`xlOpenXMLWorkbook`) alongside the original `.xlsm`. If also
   `"deleteXlsmFile" = "Y"`, deletes the original `.xlsm` file via `Scripting.FileSystemObject`.
10. **Finalize.** Hides the progress form, sets `PivotOrbConfig!A3 = "Y"` (marks as executed),
    scrolls/selects `ReportOutput!A1`, releases object references, and restores normal Excel
    settings (`Speedup_Excel(False)`).

### Helper subs/functions used by `Start_Pivot_Generation`

- **`ValidFileName(FileName, chrs)`** — returns `False` and the first offending character if the
  file name contains any of `\ / : * ? < > | [ ] "`.
- **`Speedup_Excel(BoolVal)`** — toggles `ScreenUpdating`, `DisplayAlerts`,
  `Calculation`(manual/automatic), `EnableEvents`.
- **`InitUFProgressBarBar`** — resets the `Progress` form's bar width/caption and shows it
  modeless.
- **`ProgressBar_Chart(iStep)`** — sets the progress bar width and "`n`% Complete" caption for a
  given percentage (0–100), forcing `DoEvents` so the UI updates.
- **`ORB_SheetExists(WorksheetName)`** — returns whether a sheet with that name exists.

### `RawSheetFormat(sht, JsonDict, rng)`

Reads JSON key `"rawData"` — a collection of dictionaries each with `columnName`,
`numberFormat`, `dataType`. For every entry, calls `Format_RawSheet` to apply that number format
to the corresponding column of `RawData`.

### `Format_RawSheet(sht, rg, ColName, FormatStr, dataType)`

- Finds the column in `RawData` whose header (row 1, within `A1:ZZ1`) matches `ColName`.
- Creates a temporary named range `orbformat` over that column's data rows and coerces blank
  cells to `""` and non-blank cells to numeric (`=IF(LEN(orbformat)=0,"",orbformat*1)`) — i.e.
  forces text-as-number conversion for the column.
- Applies the number format string, escaping `/` (as `\/;@`) unless the format contains
  `AM/PM` (time format).
- Deletes the temporary defined name afterward.

### `Generating_Pivots(JsonDict, PivotSheet, rng)`

Builds/rebuilds the actual pivot table (this is the core of the report):

1. Activates `PivotSheet` (`ReportOutput`) and resets the module-level `fieldMap` dictionary
   (`Public fieldMap As Scripting.Dictionary` — maps original column names to the name used as a
   pivot field, e.g. data fields get a trailing space to disambiguate from the raw field of the
   same name).
2. Creates a new `PivotCache` from the `RawData` range and, if a pivot table named `OrbitPivot`
   already exists, unprotects/unrestricts it and clears its range before rebuilding.
3. Adds a new `PivotTable` named **`OrbitPivot`** below any existing content on `ReportOutput`.
4. **`pivotLayout`** JSON key → sets row layout: `COMPACT` / `OUTLINE` / `TABULAR`
   (`xlCompactRow`/`xlOutlineRow`/`xlTabularRow`).
5. **`pageFilters`** (collection of field names) → adds each as a Page/Report filter
   (`xlPageField`), added in reverse order.
6. **`rowLabels`** (collection of `{columnName, sortType, ...}`) → adds each as a Row field
   (`xlRowField`); `sortType = "DESC"` applies descending auto-sort on that field.
7. **`columnLabels`** (same shape) → adds each as a Column field (`xlColumnField`), with the
   same optional descending sort.
8. **`dataLabels`** (collection of value-field specs) → for each entry:
   - Builds a number-format string from `numberFormat` + `applyNumberFormat` ("Y" applies it) +
     `showNegativeAsRedColor` ("Y" wraps negatives in `[Red](...)`, otherwise prefixes with `-`),
     handling both `#...` numeric formats and `%` percentage formats (percent formats get the
     `%` escaped as `\%` unless already escaped).
   - If `calculatedField` is non-empty, adds it as an Excel **calculated field**
     (`PT.CalculatedFields.Add`) instead of a plain data field; otherwise adds the raw column as
     a normal `AddDataField` with the chosen `aggregator` (mapped via `XlTypeFnc`: SUM, COUNT,
     COUNT_DISTINCT, AVG/AVERAGE, MAX, MIN, PRODUCT, "COUNT NUMBERS", STDDEV, STDDEVP,
     VAR/VARIANCE, VARP/VARIANCEP — default SUM).
   - Field name in the pivot gets a trailing space (`FieldName & " "`) to avoid clashing with a
     row/column field of the same base name; recorded in `fieldMap`.
   - **`specialCalculation`** applies a "Show Values As" transform: `PERCENTAGEOFGRANDTOTAL`,
     `PERCENTAGEOFCOLUMNTOTAL`, `PERCENTAGEOFROWTOTAL`, `PERCENTAGEOFPARENTROWTOTAL`/`...COUNT`,
     `PERCENTAGEOFPARENTCOLUMNTOTAL`/`...COUNT`, `PERCENTAGEOF` (needs `baseField`/`baseItem`),
     `PERCENTAGEOFPARENTTOTAL` (needs `baseField`). When a percentage calculation is active, the
     number format is adjusted to append `%` (or default to `"0.00 %;-0.00 %"` if none set).
   - If workbook-level `showZeroAsBlank = "Y"`, appends `;;@` to the format so zeros display as
     blank.
   - Sets the field's `numberFormat`.
   - `hideColumn = "YES"` hides the field's entire column; otherwise sets `columnWidth` from
     `columnWidth` value.
9. **`pivotCalsItems`** (collection of `{columnName, specialCalculation, baseField, numberFormat,
   ...}`) → adds Excel **calculated items** (`PivotFields(baseField).CalculatedItems.Add`) with
   the same red-negative/percent number-format logic as data labels, and applies the number
   format to the new item's field.
10. **Totals** — `Pivot_Subtotals` (see below) then `FormatPivotRowsAndColumns` (row/column
    label number formats + per-field subtotal hiding).
11. **Conditional formatting** — if `pivotConditionalFormatting = "Y"`: for every `rowLabels`
    entry and every `dataLabels` entry that has a `conditionalFormat` key (single rule dict, or
    collection of rule dicts), calls `CF.PivotCF` to add the corresponding Excel conditional
    format to that pivot field (row fields passed as `IsPivotRowField = True`, data fields as
    `False`, using the trailing-space field name).
12. **Filters** — `PivotRowsAndColumnsFilters` (see below): applies label filters
    (`pivotLabelFilter = "Y"`) and value filters (`pivotValueFilter = "Y"`) defined per field via
    `PivotFilters.PivotFilters`.
13. **Auto expand/collapse** — if `autoExpandALL = "true"`:
    - `noOfRowGroupsExpand = "-1"` → expand every row field (`ExpandCollapseDetails(..., True)`).
    - `noOfRowGroupsExpand = "0"` → collapse everything (`ExpandCollapseDetails(..., False)`).
    - Any other value `N` → collapse everything, then expand row fields whose `Position <= N`
      (`ExpandCollapseFieldDetails`).
14. **`nullDisplay`** → sets `PT.NullString` to the given text for blank/null values.
15. **`pivotProtection`** → `"Y"`: locks the pivot down (`RestrictPivotTable` — disables
    drilldown, field list, field dialog, cache refresh, and all field drag operations) and
    protects the `ReportOutput` sheet with the hard-coded password **`Orb1t!@#`**
    (`Protecting_Sheet`, allowing pivot-table use, row/column formatting). Any other value calls
    `AllowPivotTable` to re-enable all of the above.

### Supporting pivot helpers

- **`ExpandCollapseDetails(OrbPivot, ShowHide)`** — sets `ShowDetail` on every row field.
- **`ExpandCollapseFieldDetails(OrbPivot, loopint)`** — collapses everything, then expands row
  fields at positions `1..loopint`.
- **`FormatPivotRowsAndColumns` / `FormatPivotRowsAndColumns_Extend`** — for each `rowLabels`/
  `columnLabels` entry with `applyNumberFormat = "Y"`, sets the pivot field's number format
  (replacing `EEEE`→`dddd` for day names; escaping `/` unless the format is an `AM/PM` time
  format). If `pivotSubTotals <> "HIDE"` and the entry's `showTotal = "HIDE"`, calls
  `HideSubTotals` for that specific field.
- **`PivotRowsAndColumnsFilters` / `PivotRowsAndColumnsFilters_Extend`** — for each row/column
  label entry, if the workbook enables `pivotLabelFilter`/`pivotValueFilter` and the entry has a
  `labelFilters`/`valueFilters` dictionary, calls `PivotFilters.PivotFilters` for that field.
- **`XlTypeFnc(fType)`** — maps an aggregator name string to the Excel `xlConsolidationFunction`
  constant (SUM, COUNT, COUNT_DISTINCT, AVG/AVERAGE, MAX, MIN, PRODUCT, COUNT NUMBERS, STDDEV,
  STDDEVP, VAR/VARIANCE, VARP/VARIANCEP; default SUM).
- **`Pivot_Subtotals(JsonDict, OrbitPivot)`** — applies:
  - `pivotSubTotals`: `FIRST` (turn on, place at top), `LAST` (turn on, place at bottom), `HIDE`
    (turn off) — via `TurnOnTotals`/`TurnOffTotals` which set every field's `Subtotals` array.
  - `pivotGrandTotals`: `FIRST`/`LAST` → turn on both row & column grand totals; `HIDE` → turn
    off both.
  - `pivotRowGrandTotals` / `pivotColumnGrandTotals`: independently toggle row-only / column-only
    grand totals (`FIRST`/`LAST` = on, `HIDE` = off).
- **`TurnOnTotals` / `TurnOffTotals`** — set `Subtotals` on every pivot field to
  all-`True`(sum only)/all-`False`.
- **`HideSubTotals(OrbPivot, Fld)`** — turns off subtotals for one named field only.
- **`Protecting_Sheet` / `UnProtecting_Sheet`** — protect/unprotect a sheet with password
  `Orb1t!@#` (protect also permits formatting rows/columns and using pivot tables).
- **`CheckIfSheetIsProtected(sht)`** — returns `sht.ProtectContents`.
- **`RestrictPivotTable` / `AllowPivotTable`** — disable/enable drilldown, field list, field
  dialog, cache refresh, and drag-to-page/row/column/data/hide on every field of the pivot.

---

## `CF` module — conditional formatting

### `PivotCF(PT, JsonDict, PFFieldName, IsPivotRowField)`

Builds a dynamic Excel **formula-based conditional format** rule for one pivot field:

- Reads `operator`, `columnName`, `dataType`, `fontBold`, `fontItalic`, `fontUnderline`,
  `fontForeColor`, `cellFillColor`, `value1`, `value2` from the rule dictionary. Exits
  immediately if no operator is set.
- Maps `dataType` to `DATE` / `STRING` / `NUMBER` buckets.
- Optionally resolves `columnName` through the `fieldMap` to find another pivot field's data
  range to compare against (used for cross-field conditional rules); falls back to the target
  field's own range if that lookup fails.
- Builds a formula string per operator: `XLBETWEEN`, `XLNOTBETWEEN`, `XLEQUAL`, `XLNOTEQUAL`,
  `XLGREATER`, `XLGREATEREQUAL`, `XLLESS`, `XLLESSEQUAL`, `XLCONTAINS`, `XLDOESNOTCONTAIN`,
  `XLBEGINSWITH`, `XLDOESNOTBEGINSWITH`, `XLENDSWITH`, `XLDOESNOTENDSWITH`, `XLISEMPTY`,
  `XLISNOTEMPTY`. Row-field formulas additionally exclude rows containing the literal `"Total"`
  (Grand Total / Subtotal rows) via a `ISERROR(SEARCH("Total", ...))` guard, and branch by data
  type (date comparisons use `DATEVALUE`, numeric comparisons guard with `ISNUMBER`).
- Adds the formula as an `xlExpression` FormatCondition over the field's row-header column
  (`A:<field column>`), then applies font color, fill color, bold, italic, underline as
  specified.

### `GetLongFromHex(hexColor)` (private)

Converts a `#RRGGBB` (or `RRGGBB`) hex color string into the `Long` BGR-ordered value Excel's
`.Font.Color`/`.Interior.Color` expect, via `WorksheetFunction.Hex2Dec` on the reversed byte
order (`B & G & R`).

---

## `PivotFilters` module — label/value filters

### `PivotFilters(PT, JsonDict, PFFieldName, FilterType)`

Applies a **Pivot Label Filter** (`FilterType = "label"`) or **Pivot Value Filter**
(`FilterType = "value"`) to one field, based on `operator`, `dataType`, `value1`, `value2` (and,
for value filters, `columnName` naming the data field to filter by).

- Coerces `value1`/`value2` to the right VBA type for the `dataType` (`DateSerial` for
  date/time/datetime; `CDec`/`CDbl`/`CInt` for decimal/double/other numeric types), handling the
  `BETWEEN`-style operators specially (converts both values).
- **Label filters** (`ClearAllFilters` first): `XLEQUAL` → `xlCaptionEquals`, `XLNOTEQUAL` →
  `xlCaptionDoesNotEqual`, `XLBEGINSWITH`/`XLDOESNOTBEGINSWITH`, `XLENDSWITH`/
  `XLDOESNOTENDSWITH`, `XLCONTAINS`/`XLDOESNOTCONTAIN`, `XLGREATER`/`XLGREATEREQUAL`/`XLLESS`/
  `XLLESSEQUAL`, `XLBETWEEN`/`XLNOTBETWEEN` — added via `PivotFilters.Add2`.
- **Value filters** (`ClearAllFilters` first, filtered against `PT.PivotFields(columnName & " ")`
  as the data field): `XLEQUAL`/`XLNOTEQUAL`, `XLGREATER`/`XLGREATEREQUAL`/`XLLESS`/
  `XLLESSEQUAL`, `XLBETWEEN`/`XLNOTBETWEEN`, plus **Top N** style filters:
  `XLTOPCOUNT` (`xlTopCount`), `XLTOPPERCENT` (`xlTopPercent`), `XLTOPSUM` (`xlTopSum`).

---

## `JsonConverter` module — third-party library

Bundled, unmodified **VBA-JSON v2.3.1** (Tim Hall, MIT License, github.com/VBA-tools/VBA-JSON),
plus its companion **VBA-UTC v1.0.6** date-conversion helpers. Provides:

- `ParseJSON(JsonString) As Object` — parses a JSON string into nested `Scripting.Dictionary`
  (objects) / `Collection` (arrays) structures. This is what turns `PivotOrbConfig!A1` into the
  `Dict` object consumed throughout `MainModule`, `CF`, and `PivotFilters`.
- `ConvertToJson(JsonValue, ...) As String` — serializes a Dictionary/Collection/array back to
  JSON (pretty-print optional). Not used by this workbook's own logic, but available.
- `ParseUtc` / `ConvertToUtc` / `ParseIso` / `ConvertToIso` — UTC ⇄ local time and ISO‑8601 string
  conversions (used internally by `ConvertToJson` when serializing `Date` values).

This module is not workbook-specific business logic — it is a vendored open-source dependency.

---

## `Progress` UserForm

A modeless progress dialog with four named controls, manipulated entirely from `MainModule`:

- `.Border` — a static reference frame whose `.Width` is used as 100%-width baseline.
- `.Bar` — a shape/rectangle whose `.Width` is scaled to the current percentage.
- `.Text` — caption showing `"<n>% Complete"`.
- `.lblReport` — caption showing the current step's description (e.g. "Validating File Name",
  "Pivot Generation Started..", etc.).

It has no VBA code of its own — `InitUFProgressBarBar` shows it, `ProgressBar_Chart` updates it,
and `Start_Pivot_Generation` unloads it when done or when aborting early.

---

## End-to-end flow summary

```
User opens/activates "ReportOutput" sheet
        │
        ▼
Sheet1.Worksheet_Activate
  (only if PivotOrbConfig!A3 <> "Y")
        │
        ▼
MainModule.Start_Pivot_Generation
  ├─ Speedup_Excel(True) + show Progress form
  ├─ Validate file name, RawData, PivotOrbConfig, ReportOutput sheets
  ├─ Abort if already executed (A3 = "Y")
  ├─ Parse JSON from PivotOrbConfig!A1  →  Dict
  ├─ RawSheetFormat(Dict)               → number-formats RawData columns
  ├─ Generating_Pivots(Dict)            → builds "OrbitPivot" PivotTable:
  │     layout → filters → row/col labels → data fields/calculated fields
  │     → calculated items → subtotals/grand totals → number formats
  │     → conditional formatting (CF.PivotCF) → label/value filters
  │     (PivotFilters.PivotFilters) → expand/collapse → null display
  │     → protect/restrict pivot & sheet (optional)
  ├─ Optional: Save workbook as .xlsx (+ optionally delete the .xlsm)
  ├─ Set PivotOrbConfig!A3 = "Y"  (idempotency flag)
  ├─ Navigate to ReportOutput!A1
  └─ Speedup_Excel(False), unload Progress form
```

## Notable configuration keys read from `PivotOrbConfig!A1` JSON

`rawData`, `pivotLayout`, `pageFilters`, `rowLabels`, `columnLabels`, `dataLabels`,
`pivotCalsItems`, `pivotSubTotals`, `pivotGrandTotals`, `pivotRowGrandTotals`,
`pivotColumnGrandTotals`, `showZeroAsBlank`, `pivotConditionalFormatting`, `pivotLabelFilter`,
`pivotValueFilter`, `autoExpandALL`, `noOfRowGroupsExpand`, `nullDisplay`, `pivotProtection`,
`saveXlsmToXlsx`, `deleteXlsmFile`.

## Security note

The pivot-table/sheet protection password is **hard-coded** in `MainModule` as `"Orb1t!@#"`
(`Protecting_Sheet` / `UnProtecting_Sheet`). Anyone with access to the VBA project can read or
change it.
