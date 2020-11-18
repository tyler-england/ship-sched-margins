Attribute VB_Name = "Module2"
Option Explicit

Sub UpdateAkron()
    Dim bProdLinesImpt As Boolean
    bProdLinesImpt = True
    On Error GoTo errhandler
    Call UpdateMargins(bProdLinesImpt)
errhandler:
End Sub

Sub UpdateCP()
    Dim bProdLinesImpt As Boolean
    bProdLinesImpt = False
    On Error GoTo errhandler
    Call UpdateMargins(bProdLinesImpt)
errhandler:
End Sub

Function UpdateMargins(bPLsImpt As Boolean)
    Dim wbShipSched As Workbook, wsShipSched As Worksheet, wbMargin As Workbook, wsMargin As Worksheet
    Dim sMisc As String, sArrProdLines() As String, sArrResults() As String
    Dim i As Integer, j As Integer, x As Integer, y As Integer, iArrMax As Integer
    Dim iColWH As Integer, iColCO As Integer, iColPL As Integer, iColMargin As Integer
    Dim sMarginOld As Single, sMarginNew As Single
    Dim varVar As Variant, varVar2 As Variant, varData As Variant
    Dim bError As Boolean, bCont As Boolean, wbResults As Workbook
    
    ThisWorkbook.Activate
    Sheet1.Activate
    bError = True 'find margin wb
    For i = 6 To 25
        If Range("A" & i).Value > 0 Then
            sMisc = Range("B" & i).Value
            Exit For
        End If
    Next
    If sMisc > "" Then
        For Each varVar In Application.Workbooks
            If varVar.Name Like sMisc Then
                Set wbMargin = varVar
                bError = False
                Exit For
            End If
        Next
    End If
    If bError Then
        Range("A1").Select
        MsgBox "Error determining which workbook to reference for the Margins"
        Exit Function
    End If
    i = WorksheetFunction.Match("Ship*", Range("1:1"), 0) 'find ship schedule
    sMisc = Cells(1, i + 1).Value 'hopefully path is here
    If sMisc <> "" Then 'see if it's open
        varVar2 = "" 'to find name of workbook
        If InStr(sMisc, "/") > 0 Then 'prob on OneDrive
            varVar2 = "/"
        ElseIf InStr(sMisc, "\") > 0 Then 'local shared directory
            varVar2 = "\"
        End If
        varVar2 = Right(sMisc, Len(sMisc) - InStrRev(sMisc, varVar2)) 'varvar2=wb name
        For Each varVar In Application.Workbooks
            If varVar.Name Like varVar2 Then
                Set wbShipSched = varVar
                bError = False
                Exit For
            End If
        Next
    End If
    If wbShipSched Is Nothing Then 'user dialog
        If Dir$(sMisc) <> "" Then
            On Error Resume Next
                Set wbShipSched = Workbooks.Open(sMisc)
            On Error GoTo errhandler
        End If
        Do While wbShipSched Is Nothing
            varVar = Application.GetOpenFilename(, , "Choose Ship Schedule File", , False)
            If varVar = False Then Exit Function
            If Dir(varVar) <> "" Then
                Set wbShipSched = Workbooks.Open(varVar)
                If i > 0 Then 'update wb location
                    ThisWorkbook.Worksheets(1).Cells(1, i + 1).Value = wbShipSched.FullName
                End If
            End If
        Loop
    End If
    i = Year(Date) 'find FY (for ship sched sheet)
    If Month(Date) > 9 Then i = Year(Date) + 1 'FY
    For Each varVar In wbShipSched.Worksheets 'find proper worksheet
        If UCase(varVar.Name) Like CStr(i) & "*" Then
            Set wsShipSched = varVar
            Exit For
        End If
    Next
    For Each varVar In wbMargin.Worksheets 'find margin sheet
        If InStr(UCase(varVar.Name), "DETAIL") > 0 Then 'check month?
            If InStr(UCase(varVar.Name), UCase(Format(Date, "mmmm"))) > 0 Then 'month is correct
                Set wsMargin = varVar
                Exit For
            Else 'month name doesn't appear
                wbMargin.Activate
                varVar.Activate
                varVar2 = MsgBox("Do you want to use the margins on sheet '" & varVar.Name & "'?", vbYesNo)
                If varVar2 = vbYes Then
                    Set wsMargin = varVar
                    Exit For
                End If
            End If
        End If
    Next
    
    ReDim sArrProdLines(0) 'designate PL set
    sArrProdLines(0) = "REPAIR PARTS" 'going to add more 'product lines' in the future
    iArrMax = 8000 'take in 8k rows
    varData = wsMargin.Range("A1:T" & iArrMax).Value2 'find proper data on margin wb (copy in as array)
    If bPLsImpt Then 'reference Akron WB, check for sheet rearrangement
        iColWH = FindCol("Warehouse", 1, wbShipSched, 1) 'column A has warehouse
        iColCO = FindCol("Order", 2, wbShipSched, 1) 'column B has the CO
        iColPL = FindCol("Prod*Line*", 5, wbShipSched, 1) 'column E has the product line
        iColMargin = FindCol("Margin", 15, wbShipSched, 1) 'column O has the margin
    Else 'reference CLW WB, check for sheet rearrangement
        iColWH = FindCol("Whs", 4, wbShipSched, 3) 'column D has warehouse
        iColCO = FindCol("Order", 3, wbShipSched, 3) 'column C has the CO
        iColPL = FindCol("Prod*Grp*", 9, wbShipSched, 3) 'column I has the product line, but shouldn't matter
        iColMargin = FindCol("Margin", 16, wbShipSched, 3) 'column P has the margin
    End If
    ReDim sArrResults(1, 1, 1, 1, x) 'will increment 0 -> max, store values separated by "|"
    For i = 1 To 8000 'for each "row"
        If varData(i, iColWH) = "4" Then 'warehouse 4
            If bPLsImpt Then 'check product lines
                bCont = False 'default -> not a valid product line
                For j = 0 To UBound(sArrProdLines)
                    If UCase(varData(i, iColPL)) = UCase(sArrProdLines(j)) Then
                        bCont = True
                        Exit For
                    End If
                Next
            Else 'don't check product line
                bCont = True
            End If
            If bCont Then 'valid product line
                For j = 0 To x - 1 'check that CO hasn't been done already
                    If varData(i, iColCO) = sArrResults(1, 0, 0, 0, j) Then
                        bCont = False 'already considered in a previous iteration
                        Exit For
                    End If
                Next
                If bCont Then
                    sMarginNew = 0
                    y = 0 'count occurrences of CO in margin sheet/varData
                    For j = 1 To iArrMax
                        If varData(j, iColCO) = varData(i, iColCO) Then
                            y = y + 1
                        End If
                        If y > 1 Then Exit For
                    Next
                    If y > 1 Then 'more than 1 -> sum
                        For j = 1 To iArrMax
                            If varData(j, iColCO) = varData(i, iColCO) Then
                                sMarginNew = sMarginNew + varData(j, iColMargin)
                            End If
                        Next
                    Else 'that margin value
                        sMarginNew = varData(i, iColMargin)
                    End If
                    j = WorksheetFunction.CountIf(wsShipSched.Range("B:B"), varData(i, iColCO)) 'count that CO in Ship Sched
                    ReDim Preserve sArrResults(1, 1, 1, 1, x)
                    If j > 0 Then 'found at least once
                        If j = 1 Then 'appears only once
                            j = WorksheetFunction.Match(varData(i, iColCO), wsShipSched.Range("B:B"), 0)
                            sArrResults(1, 0, 0, 0, x) = varData(i, iColCO) 'add CO to report array
                            sArrResults(0, 1, 0, 0, x) = wsShipSched.Range("H" & j).Value 'add old value to report array
                            sArrResults(0, 0, 1, 0, x) = sMarginNew 'add new value to report array
                            If wsShipSched.Range("H" & j).Interior.Color <> 15773696 Then 'margin in Ship Sched isn't blue
                                With wsShipSched.Range("H" & j) 'update SS
                                    .Value = sMarginNew 'update ship sched value
                                    .Interior.Color = 15773696 'make blue
                                End With
                                sArrResults(0, 0, 0, 1, x) = "Success" 'add successful true to report array
                            Else 'ship sched already has blue margin
                                sArrResults(0, 0, 0, 1, x) = "Already in Ship Schedule"
                            End If
                        Else 'CO appears more than once in ship schedule
                            'add to failed COs [too many of this CO]?
                            sArrResults(1, 0, 0, 0, x) = varData(i, iColCO) 'add CO to report array
                            sArrResults(0, 1, 0, 0, x) = "[?]" 'add old value to report array [?]
                            sArrResults(0, 0, 1, 0, x) = sMarginNew 'add new value to report array
                            sArrResults(0, 0, 0, 1, x) = "Too many in Ship Schedule" 'add successful False
                        End If
                    Else 'CO not in ship schedule
                        'add to failed COs [not found]?
                        sArrResults(1, 0, 0, 0, x) = varData(i, iColCO) 'add CO to report array
                        sArrResults(0, 1, 0, 0, x) = "[?]" 'add old value to report array [?]
                        sArrResults(0, 0, 1, 0, x) = sMarginNew 'add new value to report array
                        sArrResults(0, 0, 0, 1, x) = "CO not in Ship Schedule" 'add successful False
                    End If
                    x = x + 1 'index for results array
                End If
            End If
        End If
    Next
    If x > 0 Then 'co report array has at least 1 element
        Set wbResults = ResultReport(sArrResults) 'create report sheet (new unsaved wb)
        wbResults.Activate 'activate report sheet
        With ActiveWindow
            If .FreezePanes Then .FreezePanes = False
            .SplitColumn = 0
            .SplitRow = 1
            .FreezePanes = True
        End With
    Else
        MsgBox "No updates were made to the Ship Schedule"
    End If
    Exit Function
errhandler:
    MsgBox "Error occurred"
    Call ErrorRep("UpdateMargins", "Function", "N/A", Err.Number, Err.Description, "")
End Function

Function FindCol(sField As String, iColOG As Integer, wbMargin As Workbook, iRowLook As Integer) As Integer
    On Error Resume Next
    FindCol = iColOG 'default is keep the OG value
    Dim i As Integer
    i = WorksheetFunction.Match(sField, wbMargin.Range(iRowLook & ":" & iRowLook), 0) 'check for field name
    If i > 0 And i <> iColOG Then FindCol = i 'reassign if field moved
End Function
Sub test()
    Dim sSavePath As String, lastF As String
    lastF = "englandt"
    sSavePath = "C:\Users\" & lastF & "\OneDrive - Barry-Wehmiller\"
        If Dir$(sSavePath & "SHIP SCHEDULE MAR*", vbDirectory) = "" Then
            MkDir sSavePath & "SHIP SCHEDULE MARGINS"
        End If
End Sub
Function ResultReport(varResults() As String) As Workbook
    Dim wbOut As Workbook, iRow As Integer, i As Integer
    Dim rngFormat As Range, dRedFont As Double, dRedBack As Double, dGreenFont As Double, dGreenBack As Double
    Dim sSavePath As String, lastF As String
    Application.ScreenUpdating = False
    lastF = Left(Application.UserName, InStr(Application.UserName, ",") - 1)
    lastF = lastF & Mid(Application.UserName, InStr(Application.UserName, ",") + 2, 1)
    '''''''hardcoded values'''''
    If Dir$("C:\users\" & lastF & "\OneDrive - Barry-Wehmiller\", vbDirectory) <> "" Then 'onedrive exists
        sSavePath = "C:\Users\" & lastF & "\OneDrive - Barry-Wehmiller\"
        If Dir$(sSavePath & "SHIP SCHEDULE MAR*", vbDirectory) = "" Then
            MkDir sSavePath & "SHIP SCHEDULE MARGINS"
        End If
        sSavePath = sSavePath & Dir$(sSavePath & "SHIP SCHEDULE MAR*", vbDirectory)
    ElseIf Dir$("C:\users\" & lastF & "\desktop\", vbDirectory) <> "" Then 'user profile exists
        If Dir$("C:\users\" & lastF & "\desktop\SHIP_SCHED*", vbDirectory) = "" Then
            MkDir "C:\users\" & lastF & "\desktop\Ship_Sched_Margins"
        End If
        sSavePath = "C:\users\" & lastF & "\desktop\Ship_Sched_Margins\"
    Else 'no one drive and/or wrong username
        If Dir$("C:\SHIP_SCHED*", vbDirectory) = "" Then
            MkDir "C:\Ship_Sched_Margins"
        End If
        sSavePath = "C:\Ship_Sched_Margins\"
    End If
    '''''''''''''''''''''''''''''
    dRedFont = 393372
    dRedBack = 13551615
    dGreenFont = 24832
    dGreenBack = 13561798
    Set wbOut = Workbooks.Add
    With wbOut.Worksheets(1)
        .Range("A1").Value = "CO"
        .Range("B1").Value = "Forecasted Margin"
        .Range("C1").Value = "Actual Margin"
        .Range("D1").Value = "Status"
        .Range("F1").Value = Format(Date, "dd mmm yyyy")
        .Range("A:D").HorizontalAlignment = xlCenter
        Columns("B:C").NumberFormat = "$#,##0.00"
        .Range("A1:F1").Font.Bold = True
        For i = 0 To UBound(varResults, 5) '5th element is the one with a uBound
            .Range("A" & i + 2).Value = varResults(1, 0, 0, 0, i)
            .Range("B" & i + 2).Value = varResults(0, 1, 0, 0, i)
            .Range("C" & i + 2).Value = varResults(0, 0, 1, 0, i)
            .Range("D" & i + 2).Value = varResults(0, 0, 0, 1, i)
            If UCase(varResults(0, 0, 0, 1, i)) = "SUCCESS" Then
                .Range("A" & i + 2 & ":D" & i + 2).Interior.Color = dGreenBack
                .Range("A" & i + 2 & ":D" & i + 2).Font.Color = dGreenFont
                If Round(varResults(0, 1, 0, 0, i), 0) <> Round(varResults(0, 0, 1, 0, i), 0) Then 'margin changed
                    .Range("B" & i + 2 & ":C" & i + 2).Interior.Color = dGreenFont
                    .Range("B" & i + 2 & ":C" & i + 2).Font.Color = 16777215
                End If
            Else
                .Range("A" & i + 2 & ":D" & i + 2).Interior.Color = dRedBack
                .Range("A" & i + 2 & ":D" & i + 2).Font.Color = dRedFont
            End If
        Next
        .Columns("A:D").AutoFit
    End With
    wbOut.SaveAs (sSavePath & "Margin_Changes_" & Format(Now, "yyyy-mm-dd_HH-mm-ss") & ".xlsx")
    Set ResultReport = wbOut
    Application.ScreenUpdating = True
End Function
