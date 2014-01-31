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
Private currentsheetname As String
Private WSFunct As Object

Public Sub CreateNewSheet(NewSheetName As String, AfterSheetName As String)
    Sheets.Add After:=Sheets(AfterSheetName)
    activesheet.name = NewSheetName
End Sub

Public Sub SetHeaders(HeaderValues, cellRange As String)
    Dim col As Integer
    Dim header As Variant
    col = 1
    For Each header In HeaderValues
        activesheet.Range(Cells(1, col), Cells(1, col)).Value = header
        col = col + 1
    Next
    activesheet.Range(cellRange).Select
    Selection.Font.Bold = True
End Sub

Public Sub MissingPcPrice(ReviewSheetName As String)
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

Public Sub PartsNotOrdered(ReviewSheetName As String)
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
    Dim i As Integer
    currentsheetname = activesheet.name
    
    Sheets("Serial File").Select
    Rows(1).Select
    Selection.Copy
    Sheets(currentsheetname).Select
    Range("A1").Select
    activesheet.Paste
    
    For i = GetLastRow("Serial File") To 2 Step -1
        Select Case Sheets("Serial File").Range("W" & i).Value
        Case "D"
            Sheets("Serial File").Select
            Rows(i).Select
            Selection.Cut
            Call PasteRows(currentsheetname, i)
        Case "I"
            Sheets("Serial File").Select
            Rows(i).Select
            Selection.Copy
            Call PasteRows(currentsheetname, i)
        End Select
    Next i
    Sheets(currentsheetname).Select
    Call TextFormatting.HeaderCorrect
    Call TextFormatting.IterateColumns
    Call SheetFormatting.AllCellsFit
End Sub
Private Sub PasteRows(sheetname As String, i As Integer)
    Sheets(sheetname).Select
    Range("A" & GetLastRow(sheetname) + 1).Select
    activesheet.Paste
    Sheets("Serial File").Select
    Rows(i).Select
    Selection.Delete Shift:=xlUp
End Sub

Private Sub AddPartNumbers(searchsheet As String, currentsheet As String)
    Dim x As Integer
    
    For x = 2 To GetLastRow(searchsheet)
        If Sheets(searchsheet).Range("A" & x).Interior.ColorIndex = 3 Then
            activesheet.Range("A" & GetLastRow(currentsheet) + 1) = Sheets(searchsheet).Range("C" & x).Value
        End If
    Next x
End Sub

Private Sub FindNotOrdered(searchsheet As String, currentsheet As String)
    Dim x As Integer
    For x = 2 To GetLastRow(searchsheet)
        If WSFunct.SumIf(Sheets(searchsheet).Range("C:C"), Sheets(searchsheet).Range("C" & x), Sheets(searchsheet).Range("G:G")) = 0 Then
            Sheets(currentsheet).Range("A" & GetLastRow(currentsheet) + 1) = Sheets(searchsheet).Range("C" & x)
        End If
    Next x
End Sub

Private Sub MatchPartNumber(searchsheet As String, currentsheet As String)
    Range("B2").FormulaR1C1 = "=VLOOKUP(RC[-1]," & searchsheet & "!R2C3:R" & GetLastRow(searchsheet) & "C4,2,False)"
    Range("B2:B" & GetLastRow(currentsheet)).Select
    Range("B2").Activate
    Selection.FillDown
End Sub

Private Function GetLastRow(name As String) As Integer
    GetLastRow = Sheets(name).Range("A36000").End(xlUp).Row
End Function

Private Sub RemoveRedundant(currentsheet As String)
    Dim i As Integer
    For i = GetLastRow(currentsheet) To 2 Step -1
        If WSFunct.CountIf(Sheets(currentsheet).Range("A:A"), Sheets(currentsheet).Range("A" & i)) > 1 Then
            Rows(i).Delete
        End If
    Next i
End Sub

