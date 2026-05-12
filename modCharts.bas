Attribute VB_Name = "modCharts"
Option Explicit

Public Sub BuildOrRefreshChart()
    Dim wsChart As Worksheet, wsSt As Worksheet
    Dim lo As ListObject
    Dim chObj As ChartObject
    Dim src As Range

    Set wsSt = ThisWorkbook.Worksheets(modConfig.STAGING_SHEET)
    On Error Resume Next
    Set lo = wsSt.ListObjects(modConfig.STAGING_LIST)
    On Error GoTo 0
    If lo Is Nothing Then
        Err.Raise vbObjectError + 701, , "未找到 Staging 表 " & modConfig.STAGING_LIST & "。"
    End If
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 702, , "Staging 表无数据行。"
    End If

    Set wsChart = GetOrCreateSheet(modConfig.CHART_SHEET)

    On Error Resume Next
    wsChart.ChartObjects(modConfig.CHART_OBJECT_NAME).Delete
    On Error GoTo 0

    ' 整张表含表头：左侧 n 列为多级分类（与透视行字段顺序一致，最右行为最内层），右侧为客户及合并列系列
    Set src = lo.Range

    Set chObj = wsChart.ChartObjects.Add(Left:=24, Top:=24, Width:=720, Height:=420)
    chObj.Name = modConfig.CHART_OBJECT_NAME

    With chObj.Chart
        .ChartType = xlColumnStacked
        .SetSourceData Source:=src, PlotBy:=xlColumns
        .HasTitle = True
        .ChartTitle.Text = modConfig.CHART_TITLE
        .HasLegend = True

        On Error Resume Next
        .Axes(xlCategory).TickLabels.Orientation = modConfig.CAT_LABEL_ANGLE
        On Error GoTo 0

        EnsureValueAxisCrossesZero chObj.Chart

        If modConfig.SHOW_DATA_LABELS Then
            .ApplyDataLabels
        Else
            Dim s As Long
            For s = 1 To .SeriesCollection.Count
                On Error Resume Next
                .SeriesCollection(s).HasDataLabels = False
                On Error GoTo 0
            Next s
        End If
    End With
End Sub

Private Sub EnsureValueAxisCrossesZero(ByVal ch As Chart)
    Dim ax As Axis
    On Error Resume Next
    Set ax = ch.Axes(xlValue)
    If ax Is Nothing Then Exit Sub

    If ax.MinimumScaleIsAuto Then
        If ax.MinimumScale > 0# Then ax.MinimumScale = 0#
    End If
    If ax.MaximumScaleIsAuto Then
        If ax.MaximumScale < 0# Then ax.MaximumScale = 0#
    End If
    On Error GoTo 0
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
