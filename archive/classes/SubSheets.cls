VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SubSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private currentsheetname$
Private WSFunct As Object

Public Sub CreateNewSheet(NewSheetName$, AfterSheetName$)
    sheets.Add After:=sheets(AfterSheetName)
    activesheet.name = NewSheetName
End Sub

Public Sub SetHeaders(HeaderValues, cellRange$)
    Dim col%
    Dim header As Variant
    col = 1
    For Each header In HeaderValues
        activesheet.Range(Cells(1, col), Cells(1, col)).Value = header
        col = col + 1
    Next
    activesheet.Range(cellRange).Select
    Selection.Font.Bold = True
End Sub

Public Sub MissingPcPrice(ReviewSheetName$)
    Set WSFunct = Application.WorksheetFunction
    currentsheetname = activesheet.name
    Call AddPartNumbers(ReviewSheetName, currentsheetname)
    If GetLastRow(currentsheetname) > 1 Then
        Call MatchPartNumber(ReviewSheetName, currentsheetname)
    End If
    Call RemoveRedundant(currentsheetname)
    Cells.HorizontalAlignment = xlCenter
    Call SheetFormatting.AllCellsFit
End Sub

Public Sub PartsNotOrdered(ReviewSheetName$)
    Set WSFunct = Application.WorksheetFunction
    currentsheetname = activesheet.name
    Call FindNotOrdered(ReviewSheetName, currentsheetname)
    If GetLastRow(currentsheetname) > 1 Then
        Call MatchPartNumber(ReviewSheetName, currentsheetname)
    End If
    Call RemoveRedundant(currentsheetname)
    Cells.HorizontalAlignment = xlCenter
    Call SheetFormatting.AllCellsFit
    SheetTools.PaneFreeze
End Sub

Public Sub MoveDeleted()
    Dim i%
    currentsheetname = activesheet.name
    
    sheets("Serial File").Select
    rows(1).Select
    Selection.Copy
    sheets(currentsheetname).Select
    Range("A1").Select
    activesheet.Paste
    
    For i = GetLastRow("Serial File") To 2 Step -1
        Select Case sheets("Serial File").Range("W" & i).Value
        Case "D"
            sheets("Serial File").Select
            rows(i).Select
            Selection.Cut
            Call PasteRows(currentsheetname, i)
        Case "I"
            sheets("Serial File").Select
            rows(i).Select
            Selection.Copy
            Call PasteRows(currentsheetname, i)
        End Select
    Next i
    sheets(currentsheetname).Select
    Call TextFormatting.HeaderCorrect
    Call TextFormatting.IterateColumns
    Call SheetFormatting.AllCellsFit
End Sub
Private Sub PasteRows(SheetName$, i%)
    sheets(SheetName).Select
    Range("A" & GetLastRow(SheetName) + 1).Select
    activesheet.Paste
    sheets("Serial File").Select
    rows(i).Select
    Selection.Delete Shift:=xlUp
End Sub

Private Sub AddPartNumbers(searchsheet$, currentsheet$)
    Dim x%
    
    For x = 2 To GetLastRow(searchsheet)
        If sheets(searchsheet).Range("A" & x).Interior.ColorIndex = 3 Then
            activesheet.Range("A" & GetLastRow(currentsheet) + 1) = sheets(searchsheet).Range("C" & x).Value
        End If
    Next x
End Sub

Private Sub FindNotOrdered(searchsheet$, currentsheet$)
    Dim x%
    For x = 2 To GetLastRow(searchsheet)
        If WSFunct.SumIf(sheets(searchsheet).Range("C:C"), sheets(searchsheet).Range("C" & x), sheets(searchsheet).Range("G:G")) = 0 Then
            sheets(currentsheet).Range("A" & GetLastRow(currentsheet) + 1) = sheets(searchsheet).Range("C" & x)
        End If
    Next x
End Sub

Private Sub MatchPartNumber(searchsheet$, currentsheet$)
    Range("B2").FormulaR1C1 = "=VLOOKUP(RC[-1]," & searchsheet & "!R2C3:R" & GetLastRow(searchsheet) & "C4,2,False)"
    Range("B2:B" & GetLastRow(currentsheet)).Select
    Range("B2").Activate
    Selection.FillDown
End Sub

Private Function GetLastRow%(name$)
    GetLastRow = sheets(name).Range("A36000").End(xlUp).Row
End Function

Private Sub RemoveRedundant(currentsheet$)
    Dim i%
    For i = GetLastRow(currentsheet) To 2 Step -1
        If WSFunct.CountIf(sheets(currentsheet).Range("A:A"), sheets(currentsheet).Range("A" & i)) > 1 Then
            rows(i).Delete
        End If
    Next i
End Sub

