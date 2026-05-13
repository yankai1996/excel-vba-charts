Attribute VB_Name = "modMain"
Option Explicit

' Entry: RefreshExistingPivot -> BuildStagingWide (target wide table) -> BuildOrRefreshChart
Public Sub Run_Pipeline()
    On Error GoTo EH
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    modPivot.RefreshExistingPivot
    modStaging.BuildStagingWide
    modCharts.BuildOrRefreshChart

    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Pipeline finished.", vbInformation
    Exit Sub

EH:
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True
    MsgBox "Pipeline failed: " & Err.Description, vbCritical, "Run_Pipeline"
End Sub
