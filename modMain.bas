Attribute VB_Name = "modMain"
Option Explicit

' 入口：清洗 -> 透视（含形状校验与刷新）-> Staging 宽表（小额合并）-> 堆叠柱图
Public Sub Run_Pipeline()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    modData.BuildCleanTable
    modPivot.EnsurePivotAndRefresh
    modStaging.BuildStagingWide
    modCharts.BuildOrRefreshChart

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "流程已完成。", vbInformation
    Exit Sub

EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "运行中断：" & Err.Description, vbCritical, "Run_Pipeline"
End Sub
