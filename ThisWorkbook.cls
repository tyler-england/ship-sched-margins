VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit
Public WithEvents MonitorApp As Application 'to monitor the application
Attribute MonitorApp.VB_VarHelpID = -1


Private Sub Workbook_Open()
    Set MonitorApp = Application
    Application.ScreenUpdating = False
End Sub

Private Sub MonitorApp_WorkbookOpen(ByVal openWB As Workbook)
    Call WorkbookList
    openWB.Activate
    Application.ScreenUpdating = True
End Sub



Private Sub MonitorApp_WorkbookBeforeClose(ByVal oldWB As Workbook, cancelClose As Boolean)
    Application.ScreenUpdating = False
    
    ThisWorkbook.Activate
    
    Dim rowNum As Integer, i As Integer
    
    Range("B6").Select
    Selection.End(xlDown).Select
    rowNum = Application.WorksheetFunction.Min(ActiveCell.Row, 150)
    If rowNum < 100 Then
        For i = 6 To rowNum
            If Range("B" & i).Value = oldWB.Name Then
                Rows(i).Delete
                Exit For
            End If
        Next i
    End If
    Range("A1").Select
    oldWB.Activate
    Application.ScreenUpdating = True
End Sub

Private Sub Workbook_AfterSave(ByVal Success As Boolean)
    If InStr(UCase(Application.UserName), "ENGLAND") > 0 Then ExportModules
End Sub
