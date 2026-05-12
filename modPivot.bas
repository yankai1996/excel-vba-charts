Attribute VB_Name = "modPivot"
Option Explicit

Public Sub EnsurePivotAndRefresh()
    On Error GoTo EH

    Dim wsClean As Worksheet, wsPiv As Worksheet
    Dim lo As ListObject
    Dim pt As PivotTable

    Set wsClean = ThisWorkbook.Worksheets(modConfig.CLEAN_SHEET)
    On Error Resume Next
    Set lo = wsClean.ListObjects(modConfig.CLEAN_LIST)
    Err.Clear
    On Error GoTo EH
    If lo Is Nothing Then
        Err.Raise vbObjectError + 501, , "未找到 " & modConfig.CLEAN_LIST & "。"
    End If
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 502, , "清洗表无数据。"
    End If

    Set wsPiv = GetOrCreateSheet(modConfig.PIVOT_SHEET)

    On Error Resume Next
    Set pt = wsPiv.PivotTables(modConfig.PIVOT_TABLE_NAME)
    Err.Clear
    On Error GoTo EH

    If Not pt Is Nothing Then
        On Error Resume Next
        pt.TableRange2.Delete
        Err.Clear
        On Error GoTo EH
    End If

    Dim pc As PivotCache
    Set pc = ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, _
        SourceData:=lo.Range)

    Set pt = pc.CreatePivotTable( _
        TableDestination:=wsPiv.Range("A3"), _
        TableName:=modConfig.PIVOT_TABLE_NAME)

    With pt
        .RowGrand = False
        .ColumnGrand = False
        On Error Resume Next
        .RowAxisLayout xlTabularRow
        .PivotFields(modConfig.COL_REGION).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        .PivotFields(modConfig.COL_YM).Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
        Err.Clear
        On Error GoTo EH

        .PivotFields(modConfig.COL_REGION).Orientation = xlRowField
        .PivotFields(modConfig.COL_REGION).Position = 1
        .PivotFields(modConfig.COL_YM).Orientation = xlRowField
        .PivotFields(modConfig.COL_YM).Position = 2
        .PivotFields(modConfig.COL_CUSTOMER).Orientation = xlColumnField
        .PivotFields(modConfig.COL_REVENUE).Orientation = xlDataField
    End With

    Dim df As PivotField
    For Each df In pt.DataFields
        df.Function = xlSum
    Next df

    pt.PivotCache.Refresh
    ValidatePivotShape pt
    Exit Sub

EH:
    Dim errNum As Long
    Dim errMsg As String
    errNum = Err.Number
    errMsg = Err.Description
    Err.Raise errNum, "modPivot.EnsurePivotAndRefresh", errMsg
End Sub

Public Sub ValidatePivotShape(ByVal pt As PivotTable)
    If pt.RowFields.Count <> 2 Then Err.Raise vbObjectError, , "透视行区必须为 2 个字段。"
    If pt.ColumnFields.Count <> 1 Then Err.Raise vbObjectError, , "透视列区必须为 1 个字段（客户）。"
    If pt.DataFields.Count <> 1 Then Err.Raise vbObjectError, , "透视值区必须为 1 个度量。"

    If StrComp(pt.RowFields(1).SourceName, modConfig.COL_REGION, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError, , "行区第 1 个字段须为 " & modConfig.COL_REGION & "（当前：" & pt.RowFields(1).SourceName & "）。"
    End If
    If StrComp(pt.RowFields(2).SourceName, modConfig.COL_YM, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError, , "行区第 2 个字段须为 " & modConfig.COL_YM & "（当前：" & pt.RowFields(2).SourceName & "）。"
    End If
    If StrComp(pt.ColumnFields(1).SourceName, modConfig.COL_CUSTOMER, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError, , "列区字段须为 " & modConfig.COL_CUSTOMER & "（当前：" & pt.ColumnFields(1).SourceName & "）。"
    End If

    Dim valField As PivotField
    Set valField = pt.DataFields(1).PivotField
    If StrComp(valField.SourceName, modConfig.COL_REVENUE, vbTextCompare) <> 0 Then
        Err.Raise vbObjectError, , "值区度量须基于字段 " & modConfig.COL_REVENUE & "（当前源字段：" & valField.SourceName & "）。"
    End If
End Sub

Private Function GetOrCreateSheet(ByVal sheetName As String) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = sheetName
    End If
    Set GetOrCreateSheet = ws
End Function
