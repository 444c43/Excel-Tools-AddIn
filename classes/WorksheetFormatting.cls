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
Public Sub SnapshotBorders()
    'border edges
    Sheets("Snapshot").Select
    
    Range("A1:K2,B4:E11,G4:J8,B25:E28,G23:J31").Select
    Call BorderEdges(xlContinuous, xlThick, xlAutomatic)
    Range("A1:K41").Select
    Call BorderEdges(xlContinuous, xlThick, xlAutomatic)
   
    'border insides
    Range("B4:E11,G4:J8,B25:E28,G23:J31").Select
    Call BorderInside(xlContinuous, xlThin, xlAutomatic)
End Sub
Public Sub MergeCells(rge As String)
    Range(rge).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
End Sub
Public Sub CellColor()
    With Selection.Interior
        .ColorIndex = 11
        .Pattern = xlSolid
    End With
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
        Range("A1:" & ActiveCell.SpecialCells(xlLastCell).Address).Select
        Call BorderEdges(xlContinuous, xlThin, xlAutomatic)
        If ActiveCell.SpecialCells(xlLastCell).Row > 1 Then
            Call BorderInside(xlContinuous, xlThin, xlAutomatic)
        End If
        Call SetFirstCell
End Sub
Private Sub BorderEdges(Style As String, Wght As String, ColIndex As Variant)
    Dim BorderObject(3) As Object
    Dim x%
    'set objects
    
    Set BorderObject(0) = Selection.Borders(xlEdgeLeft)
    Set BorderObject(1) = Selection.Borders(xlEdgeRight)
    Set BorderObject(2) = Selection.Borders(xlEdgeTop)
    Set BorderObject(3) = Selection.Borders(xlEdgeBottom)
    
    'set selected range
    For x = 0 To 3
        With BorderObject(x)
            .LineStyle = Style
            .Weight = Wght
            .ColorIndex = ColIndex
        End With
    Next x
End Sub
Private Sub BorderInside(Style As String, Wght As String, ColIndex As Variant)
    Dim BorderObject(1) As Object
    Dim x%
    
    Set BorderObject(0) = Selection.Borders(xlInsideVertical)
    Set BorderObject(1) = Selection.Borders(xlInsideHorizontal)
    'set selected range
    For x = 0 To 1
        With BorderObject(x)
            .LineStyle = Style
            .Weight = Wght
            .ColorIndex = ColIndex
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

Sub SetColumnAndRowSizes(col_rge As Variant, col_width As Variant, row_rge As Variant, row_height As Variant)
    Call SetColumnWidth(col_rge, col_width)
    Call SetRowHeight(row_rge, row_height)
End Sub

Private Sub SetColumnWidth(col_rge As Variant, col_width As Variant)
    Dim i As Byte
    
    For i = 0 To UBound(col_rge)
        Range(col_rge(i)).ColumnWidth = col_width(i)
    Next i
End Sub

Private Sub SetRowHeight(row_rge As Variant, row_height As Variant)
    Dim i As Byte
    
    For i = 0 To UBound(row_rge)
        Range(row_rge(i)).RowHeight = row_height(i)
    Next i
End Sub

