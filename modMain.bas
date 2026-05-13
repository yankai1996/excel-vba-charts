Attribute VB_Name = "modMain"
Option Explicit

' 入口：刷新/校验用户自建透视 -> Staging 宽表（多行标签 + 可选小额合并）-> 同页堆叠柱图（嵌入 ChartObject）
Public Sub Run_Pipeline()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    modPivot.RefreshExistingPivot
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
