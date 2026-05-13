Attribute VB_Name = "modConfig"
Option Explicit

'--- Sheets & ListObjects (defaults are English; set to match your workbook, e.g. Chinese names) ---
Public Const PIVOT_SHEET As String = "Pivot"
Public Const PIVOT_TABLE_NAME As String = "ptSales"

Public Const TARGET_SHEET As String = "Target"
Public Const TARGET_LIST As String = "tblTarget"

' Embedded chart floats on TARGET_SHEET next to tblTarget
Public Const CHART_OBJECT_NAME As String = "chtTargetStack"

'--- Pivot layout mode ---
' False: N row fields + exactly 1 column field (series) + 1 Sum value field.
' True:  N row fields (no column fields) + 1 Sum value; last row field = series, first N-1 = categories (requires RowFields.Count >= 2).
Public Const STACK_SERIES_FROM_LAST_ROW As Boolean = False

'--- Pivot field SourceNames (optional strict match) ---
' Leave empty to accept any column / value field SourceName (discovered at runtime).
' If non-empty, ValidatePivotShape checks they match the pivot (case-insensitive).
Public Const COL_CUSTOMER As String = ""
Public Const COL_REVENUE As String = ""

' Label used for blank pivot captions / empty row labels (must not collide with a real customer name)
Public Const BLANK_LABEL As String = "Other"

'--- Merge rules ---
Public Const ENABLE_CUSTOMER_MERGE As Boolean = True
Public Const MERGE_PROTECT_TOP_N As Long = 10
' Tie-break when |value| ties: name ascending (see MERGE_TIE_BREAK)
Public Const MERGE_TIE_BREAK As String = "NameAsc"
Public Const MERGE_ABS_THRESHOLD As Double = 500#
Public Const MERGE_PCT As Double = 0.01   ' e.g. 1%

Public Const MERGE_DENOM_ABSSUM As String = "AbsSum"
Public Const MERGE_DENOM_NETABS As String = "NetAbs"
Public Const MERGE_DENOM_POSSUM As String = "PosSum"
Public Const MERGE_PCT_DENOM As String = "AbsSum"

Public Const MERGE_COL_CAPTION As String = "Other (small amounts)"

'--- Chart ---
Public Const CHART_TITLE As String = "Revenue (stacked)"
Public Const SHOW_DATA_LABELS As Boolean = False
Public Const CAT_LABEL_ANGLE As Double = -45
