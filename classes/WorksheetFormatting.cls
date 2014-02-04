VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "WorksheetFormatting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub FormatAllWorksheets(customer_name$, acct_number$, review_period$)
    Call AllWorksheets
    Call ReviewWorksheet(customer_name, acct_number, review_period)
End Sub

Private Sub AllWorksheets()
    Dim current_sheet As Worksheet
    
    For Each current_sheet In ActiveWorkbook.Sheets
        Call IterateSheets(current_sheet.name)
    Next current_sheet
End Sub
Private Sub IterateSheets(sheet_name$)
        Sheets(sheet_name).Select
        Rows("1:1").Font.Bold = True
        Cells.EntireColumn.AutoFit
        Call BorderEdges
        Call BorderInside
        Call SetFirstCell
End Sub
Private Sub BorderEdges()
    Dim BorderObject(3) As Object
    Dim x%
    'set objects
    Range("A1:" & ActiveCell.SpecialCells(xlLastCell).Address).Select
    
    Set BorderObject(0) = Selection.Borders(xlEdgeLeft)
    Set BorderObject(1) = Selection.Borders(xlEdgeRight)
    Set BorderObject(2) = Selection.Borders(xlEdgeTop)
    Set BorderObject(3) = Selection.Borders(xlEdgeBottom)
    
    'set selected range
    For x = 0 To 3
        With BorderObject(x)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next x
End Sub
Private Sub BorderInside()
    Dim BorderObject(1) As Object
    Dim x%
    'set objects
    Range("A1:" & ActiveCell.SpecialCells(xlLastCell).Address).Select
    
    Set BorderObject(0) = Selection.Borders(xlInsideVertical)
    Set BorderObject(1) = Selection.Borders(xlInsideHorizontal)
    'set selected range
    For x = 0 To 1
        With BorderObject(x)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
    Next x
End Sub
Private Sub SetFirstCell()
    Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub
Private Sub ReviewWorksheet(customer_name$, acct_number$, review_period$)
    Sheets("Serial File").name = acct_number
    Sheets(acct_number).PageSetup.LeftHeader = customer_name
    Sheets(acct_number).PageSetup.RightHeader = review_period
End Sub
