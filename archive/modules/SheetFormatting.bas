Attribute VB_Name = "SheetFormatting"
Option Explicit

Sub MergeArray(ary As Variant)
    Dim element As Variant
    
    For Each element In ary
        Call MergeCells(CStr(element))
    Next element
End Sub

Sub MergeCells(rge$)
    Range(rge).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .MergeCells = True
    End With
End Sub
Sub CellColor(range_value$, color_value%)
    activesheet.Range(range_value).Select
    With Selection.Interior
        .ColorIndex = color_value
        .Pattern = xlSolid
    End With
End Sub
Public Sub AllCellsFit()
    Cells.EntireColumn.AutoFit
    Range("A1").Select
End Sub
Sub BorderEdges(rge$, Style$, Wght$, ColIndex As Variant)
    'changes borders of any selected range to desired linestyle and weight
    'declare variables
    Dim BorderObject(3) As Object
    Dim x%
    'set objects
    Range(rge).Select
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
Sub BorderInside(rge$, Style$, Wght$, ColIndex As Variant)
    'changes inside of any selected range to desired linestyle and weight
    'declare variables
    Dim BorderObject(1) As Object
    Dim x%
    'set objects
    Range(rge).Select
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
Sub ColDel(rge$, Shft$)
    Range(rge).Delete Shift:=Shft
End Sub
Sub ColAdd(rge$, Shft$)
    'move columns (bin system)
    Range(rge).Insert Shift:=Shft
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
