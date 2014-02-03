VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotOrdered"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AddHeaders()
    Sheets("Parts Not Ordered").Range("A1").value = "Customer Part #"
    Sheets("Parts Not Ordered").Range("B1").value = "GFC Part #"
End Sub

Public Sub FindNotOrderedParts()
    Dim i%, last_row&
    
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    For i = 2 To last_row
        'if countif(search SF column for item) = 1 and sumif( review data column for item) = 0 then
        If IsLastItemInColumn(i) = 1 Then
            If ReturnReviewSum(Sheets("Serial File").Range("E" & i)) = 0 Then
                Sheets("Parts Not Ordered").Range("A" & GetLastRow) = Sheets("Serial File").Range("E" & i)
            End If
        End If
    Next i
End Sub

Public Sub AddGFCPartNumbers()
    Dim last_row&
    
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    Sheets("Parts Not Ordered").Range("B2").FormulaR1C1 = "=VLOOKUP(RC[-1],'Serial File'!R2C5:R" & last_row & "C6,2,FALSE)"
    
    Range("B2:B" & GetLastRow - 1).FillDown
    Range("B2:B" & GetLastRow - 1).Copy
    Range("B2").PasteSpecial Paste:=xlPasteValues
End Sub

Private Function GetLastRow%()
    GetLastRow = Sheets("Parts Not Ordered").Range("A65535").End(xlUp).Row + 1
End Function

Private Function IsLastItemInColumn%(row_num%)
    IsLastItemInColumn = Application.WorksheetFunction.CountIf(Sheets("Serial File").Range("E:E"), _
        Sheets("Serial File").Range("E" & row_num))
End Function

Private Function ReturnReviewSum&(part_number$)
    ReturnReviewSum = Application.WorksheetFunction.SumIf(Sheets("Review Data").Range("E:E"), _
        part_number, Sheets("Review Data").Range("X:X"))
End Function
