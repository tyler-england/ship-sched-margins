Attribute VB_Name = "Module1"
Option Explicit
Public arrErrorEmails() As String, iNumMsgs As Integer, bNewMsg As Boolean 'for ErrorRep

Function WorkbookList()

    Dim openWB As Workbook, rowNum As Integer
    
    Application.ScreenUpdating = False
    ThisWorkbook.Activate
    Sheet1.Activate
    
    Rows("7:25").Delete
    
    Range("B6").Value = ThisWorkbook.Name
    
    rowNum = 7
    For Each openWB In Application.Workbooks
        If openWB.Name <> ThisWorkbook.Name Then
            Range("B" & rowNum).Value = openWB.Name
            rowNum = rowNum + 1
        End If
    Next openWB
    
    If rowNum > 7 Then
        Range("A6:B6").Copy
        Range("A7:B" & rowNum - 1).PasteSpecial xlPasteFormats
    End If
    
    Application.CutCopyMode = False
    Range("A1").Select

End Function

Function ExportModules() As Boolean
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'" & wbMacro.Name & "'!ExportModules", ThisWorkbook
    If Not bOpen Then wbMacro.Close savechanges:=False
    ExportModules = True
End Function

Public Sub ErrorRep(rouName, rouType, curVal, errNum, errDesc, miscInfo)
    Dim wbMacro As Workbook, varVar As Variant, bOpen As Boolean
    bNewMsg = True 'default value
    If iNumMsgs > 0 Then 'at least one email has been generated already
        For Each varVar In arrErrorEmails 'see if there were any matches
            If UCase(varVar) = UCase(ThisWorkbook.Name & "-" & errNum) Then Exit Sub 'repeat message (this was already sent this session)
        Next
    End If
    For Each varVar In Application.Workbooks
        If UCase(varVar.Name) = "MACROBOOK.XLSM" Then
            bOpen = True
            Set wbMacro = varVar
            Exit For
        End If
    Next
    If Not bOpen Then Set wbMacro = Workbooks.Open("\\PSACLW02\HOME\SHARED\MacroBook.xlsm")
    Application.Run "'MacroBook.xlsm'!ErrorReport", rouName, rouType, curVal, errNum, errDesc, miscInfo
    If Not bOpen Then wbMacro.Close savechanges:=False
    iNumMsgs = iNumMsgs + 1
    ReDim Preserve arrErrorEmails(iNumMsgs)
    arrErrorEmails(iNumMsgs) = ThisWorkbook.Name & "-" & errNum
End Sub
