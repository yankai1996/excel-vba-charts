Attribute VB_Name = "modStaging"
Option Explicit

Private Type TPair
    nm As String
    v As Double
End Type

Public Sub BuildStagingWide()
    Dim wsPiv As Worksheet
    Dim pt As PivotTable
    Dim nRow As Long
    Dim agg As Object
    Dim allCust As Object
    Dim custList() As String
    Dim nCust As Long, j As Long
    Dim rowKeys() As String
    Dim nRows As Long, rIdx As Long
    Dim out() As Variant
    Dim ncol As Long
    Dim vals() As Double
    Dim rk As Variant
    Dim inner As Object
    Dim parts() As String
    Dim seg As Long

    Set wsPiv = ThisWorkbook.Worksheets(modConfig.PIVOT_SHEET)
    On Error Resume Next
    Set pt = wsPiv.PivotTables(modConfig.PIVOT_TABLE_NAME)
    On Error GoTo 0
    If pt Is Nothing Then
        Err.Raise vbObjectError + 601, , "PivotTable not found: " & modConfig.PIVOT_TABLE_NAME & "."
    End If
    If pt.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 602, , "PivotTable has no DataBodyRange."
    End If

    nRow = pt.RowFields.Count

    Set agg = CreateObject("Scripting.Dictionary")
    Set allCust = CreateObject("Scripting.Dictionary")

    SeedCustomersFromPivot pt, allCust
    AggregateFromPivotDataBody pt, agg, allCust, nRow

    If agg.Count = 0 Then
        Err.Raise vbObjectError + 603, , "No data aggregated from pivot."
    End If
    If allCust.Count = 0 Then
        Err.Raise vbObjectError + 604, , "No customers found from pivot column field."
    End If

    custList = SortCustomersByGlobalAbs(allCust)
    nCust = UBound(custList)
    rowKeys = SortRowKeys(agg)
    nRows = UBound(rowKeys)

    ncol = nRow + nCust + IIf(modConfig.ENABLE_CUSTOMER_MERGE, 1, 0)
    ReDim out(1 To nRows, 1 To ncol)

    For rIdx = 1 To nRows
        rk = rowKeys(rIdx)
        parts = SplitRowKey(CStr(rk), nRow)
        Set inner = agg(rk)

        ReDim vals(1 To nCust)
        For j = 1 To nCust
            If inner.Exists(custList(j)) Then vals(j) = CDbl(inner(custList(j))) Else vals(j) = 0#
        Next j

        If modConfig.ENABLE_CUSTOMER_MERGE Then
            ReDim Preserve vals(1 To nCust + 1)
            ApplyMergeForRow vals, custList
        End If

        For seg = 1 To nRow
            out(rIdx, seg) = parts(seg)
        Next seg
        For j = 1 To nCust
            out(rIdx, nRow + j) = vals(j)
        Next j
        If modConfig.ENABLE_CUSTOMER_MERGE Then
            out(rIdx, ncol) = vals(nCust + 1)
        End If
    Next rIdx

    Dim wsTgt As Worksheet, loTgt As ListObject
    Dim hdr() As String

    Set wsTgt = GetOrCreateSheet(modConfig.TARGET_SHEET)
    ReDim hdr(1 To ncol)
    For seg = 1 To nRow
        hdr(seg) = pt.RowFields(seg).SourceName
    Next seg
    For j = 1 To nCust
        hdr(nRow + j) = custList(j)
    Next j
    If modConfig.ENABLE_CUSTOMER_MERGE Then hdr(ncol) = modConfig.MERGE_COL_CAPTION

    Set loTgt = EnsureListWithHeaders(wsTgt, modConfig.TARGET_LIST, hdr)
    WriteBody loTgt, out
End Sub

Private Function SplitRowKey(ByVal rk As String, ByVal nRow As Long) As String()
    Dim parts() As String
    Dim raw() As String
    Dim i As Long
    ReDim parts(1 To nRow)
    raw = Split(rk, vbTab)
    For i = 1 To nRow
        If i - 1 + LBound(raw) <= UBound(raw) Then
            parts(i) = raw(i - 1 + LBound(raw))
        Else
            parts(i) = ""
        End If
        If Len(Trim$(parts(i))) = 0 Then parts(i) = modConfig.BLANK_LABEL
    Next i
    SplitRowKey = parts
End Function

Private Sub SeedCustomersFromPivot(ByVal pt As PivotTable, ByVal allCust As Object)
    Dim pf As PivotField
    Dim pi As PivotItem
    Dim cust As String

    On Error Resume Next
    Set pf = pt.PivotFields(modConfig.COL_CUSTOMER)
    If Err.Number <> 0 Then
        Err.Clear
        Exit Sub
    End If
    On Error GoTo 0

    For Each pi In pf.PivotItems
        On Error Resume Next
        If pi.Visible Then
            cust = NormalizePivotCaption(pi.Caption)
            If Len(cust) = 0 Then cust = modConfig.BLANK_LABEL
            If Not allCust.Exists(cust) Then allCust.Add cust, 0#
        End If
        On Error GoTo 0
    Next pi
End Sub

Private Sub AggregateFromPivotDataBody(ByVal pt As PivotTable, ByVal agg As Object, ByVal allCust As Object, ByVal nRow As Long)
    Dim c As Range
    Dim pc As PivotCell
    Dim cust As String
    Dim rk As String
    Dim inner As Object
    Dim v As Double

    For Each c In pt.DataBodyRange.Cells
        On Error Resume Next
        Set pc = c.PivotCell
        If Err.Number <> 0 Then
            Err.Clear
            GoTo NextCell
        End If
        On Error GoTo 0

        If pc.PivotCellType <> xlPivotCellValue Then GoTo NextCell
        If pc.RowItems.Count <> nRow Then GoTo NextCell
        If pc.ColumnItems.Count < 1 Then GoTo NextCell

        rk = BuildRowKeyFromPivotCell(pc, nRow)
        cust = NormalizePivotCaption(CStr(pc.ColumnItems(1).Caption))
        v = PivotCellValueToDouble(c)

        If Len(cust) = 0 Then cust = modConfig.BLANK_LABEL

        If Not agg.Exists(rk) Then agg.Add rk, CreateObject("Scripting.Dictionary")
        Set inner = agg(rk)
        If inner.Exists(cust) Then
            inner(cust) = CDbl(inner(cust)) + v
        Else
            inner.Add cust, v
        End If

        If Not allCust.Exists(cust) Then allCust.Add cust, 0#
        allCust(cust) = CDbl(allCust(cust)) + Abs(v)

NextCell:
    Next c
End Sub

Private Function BuildRowKeyFromPivotCell(ByVal pc As PivotCell, ByVal nRow As Long) As String
    Dim i As Long
    Dim s As String
    Dim t As String

    For i = 1 To nRow
        t = Trim$(CStr(pc.RowItems(i).Caption))
        If Len(t) = 0 Then t = modConfig.BLANK_LABEL
        If i > 1 Then s = s & vbTab
        s = s & t
    Next i
    BuildRowKeyFromPivotCell = s
End Function

Private Function NormalizePivotCaption(ByVal s As String) As String
    s = Trim$(s)
    If StrComp(s, "(blank)", vbTextCompare) = 0 Then
        NormalizePivotCaption = modConfig.BLANK_LABEL
    Else
        NormalizePivotCaption = s
    End If
End Function

Private Function PivotCellValueToDouble(ByVal c As Range) As Double
    On Error Resume Next
    If IsEmpty(c.Value) Then
        PivotCellValueToDouble = 0#
    ElseIf IsNumeric(c.Value) Then
        PivotCellValueToDouble = CDbl(c.Value)
    Else
        PivotCellValueToDouble = 0#
    End If
    On Error GoTo 0
End Function

Private Sub ApplyMergeForRow(ByRef vals() As Double, ByRef custList() As String)
    Dim n As Long, pairs() As TPair
    Dim i As Long, cutIdx As Long, cutoff As Double
    Dim prot() As Boolean
    Dim T As Double
    Dim mergeSum As Double
    Dim vv As Double
    Dim okPct As Boolean
    Dim idx As Long

    n = UBound(vals) - 1
    ReDim pairs(1 To n)
    For i = 1 To n
        pairs(i).nm = custList(i)
        pairs(i).v = vals(i)
    Next i

    SortPairsByAbsDesc pairs

    cutIdx = Application.WorksheetFunction.Min(modConfig.MERGE_PROTECT_TOP_N, n)
    cutoff = Abs(pairs(cutIdx).v)

    ReDim prot(1 To n)
    For i = 1 To n
        prot(i) = (Abs(pairs(i).v) >= cutoff)
    Next i

    T = ComputeTRow(vals, n)

    mergeSum = 0#
    For i = 1 To n
        If prot(i) Then GoTo NextI2
        vv = pairs(i).v
        okPct = (T > 0#) And (Abs(vv) < T * modConfig.MERGE_PCT)
        If (Abs(vv) < modConfig.MERGE_ABS_THRESHOLD) Or okPct Then
            mergeSum = mergeSum + vv
            idx = IndexOfCust(custList, pairs(i).nm)
            If idx > 0 Then vals(idx) = 0#
        End If
NextI2:
    Next i

    vals(n + 1) = mergeSum
End Sub

Private Function ComputeTRow(ByRef vals() As Double, ByVal nCust As Long) As Double
    Dim s As Double, net As Double, pos As Double
    Dim j As Long

    Select Case modConfig.MERGE_PCT_DENOM
        Case modConfig.MERGE_DENOM_NETABS
            net = 0#
            For j = 1 To nCust
                net = net + vals(j)
            Next j
            ComputeTRow = Abs(net)
        Case modConfig.MERGE_DENOM_POSSUM
            pos = 0#
            For j = 1 To nCust
                pos = pos + WorksheetFunction.Max(vals(j), 0#)
            Next j
            ComputeTRow = pos
        Case Else
            s = 0#
            For j = 1 To nCust
                s = s + Abs(vals(j))
            Next j
            ComputeTRow = s
    End Select
End Function

Private Function IndexOfCust(ByRef custList() As String, ByVal nm As String) As Long
    Dim i As Long
    For i = LBound(custList) To UBound(custList)
        If custList(i) = nm Then IndexOfCust = i: Exit Function
    Next i
    IndexOfCust = 0
End Function

Private Sub SortPairsByAbsDesc(ByRef pairs() As TPair)
    Dim i As Long, j As Long, tmp As TPair
    For i = LBound(pairs) To UBound(pairs) - 1
        For j = i + 1 To UBound(pairs)
            If Abs(pairs(j).v) > Abs(pairs(i).v) _
               Or (Abs(pairs(j).v) = Abs(pairs(i).v) And pairs(j).nm < pairs(i).nm) Then
                tmp = pairs(i): pairs(i) = pairs(j): pairs(j) = tmp
            End If
        Next j
    Next i
End Sub

Private Function SortCustomersByGlobalAbs(ByVal allCust As Object) As String()
    Dim arr() As TPair
    Dim k As Variant, n As Long, i As Long
    Dim outCust() As String

    n = allCust.Count
    ReDim arr(1 To n)
    i = 1
    For Each k In allCust.Keys
        arr(i).nm = CStr(k)
        arr(i).v = CDbl(allCust(k))
        i = i + 1
    Next k
    SortPairsByAbsDesc arr

    ReDim outCust(1 To n)
    For i = 1 To n
        outCust(i) = arr(i).nm
    Next i
    SortCustomersByGlobalAbs = outCust
End Function

Private Function SortRowKeys(ByVal agg As Object) As String()
    Dim arr() As String
    Dim k As Variant, n As Long, i As Long, j As Long
    Dim t As String

    n = agg.Count
    ReDim arr(1 To n)
    i = 1
    For Each k In agg.Keys
        arr(i) = CStr(k)
        i = i + 1
    Next k

    For i = 1 To n - 1
        For j = i + 1 To n
            If RowKeyLess(arr(j), arr(i)) Then
                t = arr(i): arr(i) = arr(j): arr(j) = t
            End If
        Next j
    Next i
    SortRowKeys = arr
End Function

Private Function RowKeyLess(ByVal a As String, ByVal b As String) As Boolean
    Dim sa() As String, sb() As String
    Dim i As Long, ubA As Long, ubB As Long, u As Long
    Dim cmp As Long

    sa = Split(a, vbTab)
    sb = Split(b, vbTab)
    ubA = UBound(sa)
    ubB = UBound(sb)
    u = Application.WorksheetFunction.Min(ubA, ubB)

    For i = LBound(sa) To u
        cmp = StrComp(sa(i), sb(i), vbTextCompare)
        If cmp < 0 Then RowKeyLess = True: Exit Function
        If cmp > 0 Then RowKeyLess = False: Exit Function
    Next i

    RowKeyLess = (ubA < ubB)
End Function

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

Private Function EnsureListWithHeaders(ByVal ws As Worksheet, ByVal loName As String, ByRef hdr() As String) As ListObject
    Dim lo As ListObject
    Dim c As Long, lastCol As Long
    lastCol = UBound(hdr)
    On Error Resume Next
    Set lo = ws.ListObjects(loName)
    On Error GoTo 0
    If Not lo Is Nothing Then lo.Delete

    For c = 1 To lastCol
        ws.Cells(1, c).Value = hdr(c)
    Next c
    Dim rng As Range
    Set rng = ws.Range(ws.Cells(1, 1), ws.Cells(1, lastCol))
    Set lo = ws.ListObjects.Add(xlSrcRange, rng, , xlYes)
    lo.Name = loName
    Set EnsureListWithHeaders = lo
End Function

Private Sub WriteBody(ByVal lo As ListObject, ByVal out As Variant)
    Dim n As Long, rng As Range
    n = UBound(out, 1)
    Do While lo.ListRows.Count > 0
        lo.ListRows(1).Delete
    Loop
    If n = 0 Then Exit Sub
    Set rng = lo.HeaderRowRange.Cells(1, 1).Resize(n + 1, UBound(out, 2))
    On Error Resume Next
    lo.Resize rng
    On Error GoTo 0
    lo.DataBodyRange.Value = out
End Sub
