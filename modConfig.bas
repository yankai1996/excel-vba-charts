Attribute VB_Name = "modConfig"
Option Explicit

'--- Sheets & ListObjects ---
Public Const PIVOT_SHEET As String = "透视"
Public Const PIVOT_TABLE_NAME As String = "ptSales"

Public Const TARGET_SHEET As String = "Target"
Public Const TARGET_LIST As String = "tblTarget"

' 嵌入图与 tblTarget 同在 TARGET_SHEET 上（悬浮于网格之上）
Public Const CHART_OBJECT_NAME As String = "chtTargetStack"

'--- Pivot field names (列/值校验；行字段名与个数由用户透视决定) ---
Public Const COL_CUSTOMER As String = "客户"
Public Const COL_REVENUE As String = "收入"

'--- Merge rules ---
Public Const ENABLE_CUSTOMER_MERGE As Boolean = True
Public Const MERGE_PROTECT_TOP_N As Long = 10
' 并列时的二次排序：当前实现固定为「客户名字典序」（与 MERGE_TIE_BREAK 语义一致）
Public Const MERGE_TIE_BREAK As String = "NameAsc"
Public Const MERGE_ABS_THRESHOLD As Double = 500#
Public Const MERGE_PCT As Double = 0.01   ' e.g. 1%

Public Const MERGE_DENOM_ABSSUM As String = "AbsSum"
Public Const MERGE_DENOM_NETABS As String = "NetAbs"
Public Const MERGE_DENOM_POSSUM As String = "PosSum"
Public Const MERGE_PCT_DENOM As String = "AbsSum"

Public Const MERGE_COL_CAPTION As String = "其他（客户小额）"

'--- Chart ---
Public Const CHART_TITLE As String = "收入（堆叠）"
Public Const SHOW_DATA_LABELS As Boolean = False
Public Const CAT_LABEL_ANGLE As Double = -45
