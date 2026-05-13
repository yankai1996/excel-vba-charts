Attribute VB_Name = "modCharts"
Option Explicit

Public Sub BuildOrRefreshChart()
    Dim wsTgt As Worksheet
    Dim lo As ListObject
    Dim chObj As ChartObject
    Dim src As Range
    Dim leftPts As Double, topPts As Double

    Set wsTgt = ThisWorkbook.Worksheets(modConfig.TARGET_SHEET)
    On Error Resume Next
    Set lo = wsTgt.ListObjects(modConfig.TARGET_LIST)
    On Error GoTo 0
    If lo Is Nothing Then
        Err.Raise vbObjectError + 701, , "未找到目标宽表 " & modConfig.TARGET_LIST & "。"
    End If
    If lo.DataBodyRange Is Nothing Then
        Err.Raise vbObjectError + 702, , "目标宽表无数据行。"
    End If

    On Error Resume Next
    wsTgt.ChartObjects(modConfig.CHART_OBJECT_NAME).Delete
    On Error GoTo 0

    ' 整张表含表头：左侧 n 列为多级分类，右侧为客户及合并列系列
    Set src = lo.Range

    leftPts = lo.Range.Left + lo.Range.Width + 20#
    topPts = lo.Range.Top
    If leftPts < 12# Then leftPts = 24#
    If topPts < 12# Then topPts = 24#

    Set chObj = wsTgt.ChartObjects.Add(Left:=leftPts, Top:=topPts, Width:=720, Height:=420)
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
