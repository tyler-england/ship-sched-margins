VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Worksheet_Change(ByVal Target As Range)

    Dim xRange As Range, rowNum As Integer, xVal As Boolean, curRange As Range, i As Integer

    Set xRange = Range("A:A")
    Set curRange = ActiveCell
    
    If Not Application.Intersect(xRange, Range(Target.Address)) Is Nothing Then
           
        Application.ScreenUpdating = False
        Range("B6").Select
        Selection.End(xlDown).Select
        rowNum = Application.WorksheetFunction.Min(ActiveCell.Row, 150)
        If rowNum < 100 Then
            For i = 6 To rowNum
                If Range("A" & i).Value > 0 Then
                    If xVal Then
                        Range("A6:A" & rowNum).ClearContents
                    Else
                        xVal = True
                    End If
                End If
            Next i
        End If
        
        curRange.Select
        Application.ScreenUpdating = True
        
    End If
End Sub
