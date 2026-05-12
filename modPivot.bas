Attribute VB_Name = "modPivot"
Option Explicit

Public Sub RefreshExistingPivot()
    On Error GoTo EH

    Dim wsPiv As Worksheet
    Dim pt As PivotTable
    Dim errNum As Long
    Dim errMsg As String

    Set wsPiv = ThisWorkbook.Worksheets(modConfig.PIVOT_SHEET)

    On Error Resume Next
    Set pt = wsPiv.PivotTables(modConfig.PIVOT_TABLE_NAME)
    Err.Clear
    On Error GoTo EH

    If pt Is Nothing Then
        Err.Raise vbObjectError + 511, , "未找到透视表 " & modConfig.PIVOT_TABLE_NAME & "（工作表 """ & modConfig.PIVOT_SHEET & """）。"
    End If
    If pt.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 512, , "透视表无数据区（DataBodyRange），请检查数据源与字段布局。"
    End If

    On Error Resume Next
    pt.PivotCache.Refresh
    If Err.Number <> 0 Then
        errNum = Err.Number
        errMsg = Err.Description
        Err.Clear
        On Error GoTo EH
        Err.Raise errNum, , "透视刷新失败：" & errMsg
    End If
    Err.Clear
    On Error GoTo EH

    ValidatePivotShape pt
    Exit Sub

EH:
    errNum = Err.Number
    errMsg = Err.Description
    Err.Raise errNum, "modPivot.RefreshExistingPivot", errMsg
End Sub

Public Sub ValidatePivotShape(ByVal pt As PivotTable)
    If pt.RowFields.Count < 1 Then
        Err.Raise vbObjectError, , "透视行区至少需要 1 个行字段。"
    End If
    If pt.ColumnFields.Count <> 1 Then
        Err.Raise vbObjectError, , "透视列区必须为 1 个字段（客户）。"
    End If
    If pt.DataFields.Count <> 1 Then
        Err.Raise vbObjectError, , "透视值区必须为 1 个度量。"
    End If

    If StrComp(pt.ColumnFields(1).SourceName, modConfig.COL_CUSTOMER, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError, , "列区字段须为 " & modConfig.COL_CUSTOMER & "（当前：" & pt.ColumnFields(1).SourceName & "）。"
    End If

    Dim valField As PivotField
    Set valField = pt.DataFields(1).PivotField
    If StrComp(valField.SourceName, modConfig.COL_REVENUE, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError, , "值区度量须基于字段 " & modConfig.COL_REVENUE & "（当前源字段：" & valField.SourceName & "）。"
    End If
    If pt.DataFields(1).Function <> xlSum Then
        Err.Raise vbObjectError, , "值区度量须为「求和」（当前聚合函数不是求和）。"
    End If
End Sub
