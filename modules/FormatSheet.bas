Attribute VB_Name = "FormatSheet"
Option Explicit

Public Sub ResizeAllCells()
    Cells.EntireColumn.AutoFit
    Range("A1").Select
End Sub

'Sub MergeCells(Rge As String)
'    Range(Rge).Select
'    With Selection
'        .HorizontalAlignment = xlCenter
'        .VerticalAlignment = xlCenter
'        .MergeCells = True
'    End With
'End Sub
'Sub CellColor()
'    With Selection.Interior
'        .ColorIndex = 11
'        .Pattern = xlSolid
'    End With
'End Sub
'
'Sub BorderEdges(Rge As String, Style As String, Wght As String, ColIndex As Variant)
'    'changes borders of any selected range to desired linestyle and weight
'    'declare variables
'    Dim BorderObject(3) As Object
'    Dim x As Integer
'    'set objects
'    Range(Rge).Select
'    Set BorderObject(0) = Selection.Borders(xlEdgeLeft)
'    Set BorderObject(1) = Selection.Borders(xlEdgeRight)
'    Set BorderObject(2) = Selection.Borders(xlEdgeTop)
'    Set BorderObject(3) = Selection.Borders(xlEdgeBottom)
'    'set selected range
'    For x = 0 To 3
'        With BorderObject(x)
'            .LineStyle = Style
'            .Weight = Wght
'            .ColorIndex = ColIndex
'        End With
'    Next x
'End Sub
'Sub BorderInside(Rge As String, Style As String, Wght As String, ColIndex As Variant)
'    'changes inside of any selected range to desired linestyle and weight
'    'declare variables
'    Dim BorderObject(1) As Object
'    Dim x As Integer
'    'set objects
'    Range(Rge).Select
'    Set BorderObject(0) = Selection.Borders(xlInsideVertical)
'    Set BorderObject(1) = Selection.Borders(xlInsideHorizontal)
'    'set selected range
'    For x = 0 To 1
'        With BorderObject(x)
'            .LineStyle = Style
'            .Weight = Wght
'            .ColorIndex = ColIndex
'        End With
'    Next x
'End Sub
'Sub ColDel(Rge As String, Shft As String)
'    Range(Rge).Delete Shift:=Shft
'End Sub
'Sub ColAdd(Rge As String, Shft As String)
'    'move columns (bin system)
'    Range(Rge).Insert Shift:=Shft
'End Sub
