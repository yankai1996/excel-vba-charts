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
        Err.Raise vbObjectError + 511, , "PivotTable not found: " & modConfig.PIVOT_TABLE_NAME & " on sheet """ & modConfig.PIVOT_SHEET & """."
    End If
    If pt.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 512, , "PivotTable has no DataBodyRange; check source data and field layout."
    End If

    On Error Resume Next
    pt.PivotCache.Refresh
    If Err.Number <> 0 Then
        errNum = Err.Number
        errMsg = Err.Description
        Err.Clear
        On Error GoTo EH
        Err.Raise errNum, , "PivotCache.Refresh failed: " & errMsg
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
        Err.Raise vbObjectError, , "Pivot row area must have at least one row field."
    End If
    If pt.DataFields.Count <> 1 Then
        Err.Raise vbObjectError, , "Pivot data area must have exactly one data field."
    End If

    Dim valField As PivotField
    Set valField = pt.DataFields(1).PivotField
    If Len(modConfig.COL_REVENUE) > 0 Then
        If StrComp(valField.SourceName, modConfig.COL_REVENUE, vbTextCompare) <> 0 Then
            Err.Raise vbObjectError, , "Data field must be based on " & modConfig.COL_REVENUE & " (found: " & valField.SourceName & ")."
        End If
    End If
    If pt.DataFields(1).Function <> xlSum Then
        Err.Raise vbObjectError, , "Data field must aggregate with Sum (current function is not Sum)."
    End If

    If modConfig.STACK_SERIES_FROM_LAST_ROW Then
        If pt.ColumnFields.Count <> 0 Then
            Err.Raise vbObjectError, , "When STACK_SERIES_FROM_LAST_ROW is True, the pivot must have no column fields (found " & pt.ColumnFields.Count & ")."
        End If
        If pt.RowFields.Count < 2 Then
            Err.Raise vbObjectError, , "When STACK_SERIES_FROM_LAST_ROW is True, the pivot must have at least two row fields (outer = category axis, innermost = stack series)."
        End If
    Else
        If pt.ColumnFields.Count <> 1 Then
            Err.Raise vbObjectError, , "Pivot column area must have exactly one field (stack series)."
        End If
        If Len(modConfig.COL_CUSTOMER) > 0 Then
            If StrComp(pt.ColumnFields(1).SourceName, modConfig.COL_CUSTOMER, vbTextCompare) <> 0 Then
                Err.Raise vbObjectError, , "Column field must be " & modConfig.COL_CUSTOMER & " (found: " & pt.ColumnFields(1).SourceName & ")."
            End If
        End If
    End If
End Sub
