VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetTools"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private LastRow&

Function CountUnique&(column_letter$)
    Dim x&, last_row&
    
    LastRow = ActiveCell.SpecialCells(xlLastCell).Row
    
    For x = 2 To LastRow
        If IterateCountIf(column_letter, x, Sheets(activesheet.name)) = 1 Then
            CountUnique = CountUnique + 1
        End If
    Next x
End Function

Private Function IterateCountIf&(column_letter$, i&, sheet_name As Object)
    Dim Func As Object
    Dim search_range$, current_cell$
    
    search_range = column_letter & i & ":" & column_letter & LastRow
    current_cell = column_letter & i
    
    Set Func = Application.WorksheetFunction
    IterateCountIf = Func.CountIf(sheet_name.Range(search_range), sheet_name.Range(current_cell))
End Function

