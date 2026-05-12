Attribute VB_Name = "modData"
Option Explicit

Public Sub BuildCleanTable()
    Dim wsSrc As Worksheet, wsClean As Worksheet
    Dim loSrc As ListObject, loClean As ListObject
    Dim rBody As Range
    Dim r As Long
    Dim hdrMap As Object
    Dim colReg As Long, colYm As Long, colCust As Long, colRev As Long
    Dim out() As Variant
    Dim nRows As Long, i As Long
    Dim sRegion As String, sYm As String, sCust As String
    Dim revVal As Double

    Set wsSrc = GetOrCreateSheet(modConfig.SRC_SHEET)
    Set wsClean = GetOrCreateSheet(modConfig.CLEAN_SHEET)

    Set loSrc = EnsureListObject(wsSrc, modConfig.SRC_LIST, Array(modConfig.COL_REGION, modConfig.COL_YM, modConfig.COL_CUSTOMER, modConfig.COL_REVENUE))
    Set loClean = EnsureListObject(wsClean, modConfig.CLEAN_LIST, Array(modConfig.COL_REGION, modConfig.COL_YM, modConfig.COL_CUSTOMER, modConfig.COL_REVENUE))

    On Error Resume Next
    Set rBody = loSrc.DataBodyRange
    On Error GoTo 0
    If rBody Is Nothing Then
        Err.Raise vbObjectError + 401, , "源表 " & modConfig.SRC_LIST & " 没有数据行。"
    End If

    Set hdrMap = BuildHeaderMap(loSrc.HeaderRowRange)
    colReg = FindCol(hdrMap, modConfig.COL_REGION)
    colYm = FindCol(hdrMap, modConfig.COL_YM)
    colCust = FindCol(hdrMap, modConfig.COL_CUSTOMER)
    colRev = FindCol(hdrMap, modConfig.COL_REVENUE)
    If colReg = 0 Or colYm = 0 Or colCust = 0 Or colRev = 0 Then
        Err.Raise vbObjectError + 402, , "源表缺少列：需包含 " & modConfig.COL_REGION & "、" & modConfig.COL_YM & "、" & modConfig.COL_CUSTOMER & "、" & modConfig.COL_REVENUE & "。"
    End If

    nRows = rBody.Rows.Count
    ReDim out(1 To nRows, 1 To 4)

    For i = 1 To nRows
        r = rBody.Row + i - 1
        sRegion = Trim$(CStr(wsSrc.Cells(r, colReg).Value2))
        sYm = Trim$(CStr(wsSrc.Cells(r, colYm).Value2))
        sCust = Trim$(CStr(wsSrc.Cells(r, colCust).Value2))
        revVal = ParseNumber(wsSrc.Cells(r, colRev))

        If Len(sRegion) = 0 Then sRegion = "其他"
        If Len(sYm) = 0 Then sYm = "其他"
        If Len(sCust) = 0 Then sCust = "其他"

        out(i, 1) = sRegion
        out(i, 2) = sYm
        out(i, 3) = sCust
        out(i, 4) = revVal
    Next i

    WriteListBody loClean, out
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

Private Function EnsureListObject(ByVal ws As Worksheet, ByVal loName As String, ByVal headers As Variant) As ListObject
    Dim lo As ListObject
    On Error Resume Next
    Set lo = ws.ListObjects(loName)
    On Error GoTo 0
    If lo Is Nothing Then
        Dim c As Long, lastCol As Long
        lastCol = UBound(headers)
        For c = 1 To lastCol
            ws.Cells(1, c).Value = headers(c)
        Next c
        Dim rng As Range
        Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
        Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
        lo.Name = loName
    End If
    Set EnsureListObject = lo
End Function

Private Function BuildHeaderMap(ByVal hdr As Range) As Object
    Dim d As Object, c As Long, t As String
    Set d = CreateObject("Scripting.Dictionary")
    For c = 1 To hdr.Columns.Count
        t = Trim$(CStr(hdr.Cells(1, c).Value2))
        If Len(t) > 0 And Not d.Exists(t) Then d.Add t, hdr.Cells(1, c).Column
    Next c
    If Not d.Exists(modConfig.COL_REGION) And d.Exists("区域") Then d.Add modConfig.COL_REGION, d("区域")
    Set BuildHeaderMap = d
End Function

Private Function FindCol(ByVal hdrMap As Object, ByVal logicalName As String) As Long
    If hdrMap.Exists(logicalName) Then FindCol = hdrMap(logicalName) Else FindCol = 0
End Function

Private Function ParseNumber(ByVal rng As Range) As Double
    Dim s As String
    On Error Resume Next
    If IsNumeric(rng.Value2) And Not IsEmpty(rng.Value2) Then
        ParseNumber = CDbl(rng.Value2)
        Exit Function
    End If
    s = Trim$(CStr(rng.Value2))
    If Len(s) = 0 Then ParseNumber = 0#: Exit Function
    s = Replace(s, ",", "")
    s = Replace(s, "￥", "")
    s = Replace(s, "$", "")
    s = Trim$(s)
    If Len(s) = 0 Then ParseNumber = 0#: Exit Function
    If IsNumeric(s) Then ParseNumber = CDbl(s) Else ParseNumber = 0#
End Function

Private Sub WriteListBody(ByVal lo As ListObject, ByVal out As Variant)
    Dim n As Long, rng As Range, ws As Worksheet
    n = UBound(out, 1)
    Set ws = lo.Parent
    Do While lo.ListRows.Count > 0
        lo.ListRows(1).Delete
    Loop
    If n = 0 Then Exit Sub
    Set rng = lo.HeaderRowRange.Resize(n + 1, UBound(out, 2))
    On Error Resume Next
    lo.Resize rng
    On Error GoTo 0
    lo.DataBodyRange.Value = out
End Sub
