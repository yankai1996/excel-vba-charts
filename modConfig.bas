Attribute VB_Name = "modConfig"
Option Explicit

'--- Sheets & ListObjects (create these names in the workbook or change here)---
Public Const SRC_SHEET As String = "源数据"
Public Const SRC_LIST As String = "tblSource"

Public Const CLEAN_SHEET As String = "清洗"
Public Const CLEAN_LIST As String = "tblClean"

Public Const PIVOT_SHEET As String = "透视"
Public Const PIVOT_TABLE_NAME As String = "ptSales"

Public Const STAGING_SHEET As String = "Staging"
Public Const STAGING_LIST As String = "tblStage"

Public Const CHART_SHEET As String = "图表"
Public Const CHART_OBJECT_NAME As String = "chtStagingStack"

'--- Logical column headers on CLEAN table (after header_map) ---
Public Const COL_REGION As String = "地区"
Public Const COL_YM As String = "年月"
Public Const COL_CUSTOMER As String = "客户"
Public Const COL_REVENUE As String = "收入"

'--- Merge rules (plan) ---
Public Const ENABLE_CUSTOMER_MERGE As Boolean = True
Public Const MERGE_PROTECT_TOP_N As Long = 10
' 并列时的二次排序：当前实现固定为「客户名字典序」（与 MERGE_TIE_BREAK 语义一致）
Public Const MERGE_TIE_BREAK As String = "NameAsc"
Public Const MERGE_ABS_THRESHOLD As Double = 500#
Public Const MERGE_PCT As Double = 0.01   ' e.g. 1%

' MERGE_PCT_DENOM: use one of MERGE_DENOM_*
Public Const MERGE_DENOM_ABSSUM As String = "AbsSum"
Public Const MERGE_DENOM_NETABS As String = "NetAbs"
Public Const MERGE_DENOM_POSSUM As String = "PosSum"
Public Const MERGE_PCT_DENOM As String = "AbsSum"

Public Const MERGE_COL_CAPTION As String = "其他（客户小额）"

'--- Chart ---
Public Const CHART_TITLE As String = "收入（堆叠）"
Public Const SHOW_DATA_LABELS As Boolean = False
Public Const CAT_LABEL_ANGLE As Double = -45
