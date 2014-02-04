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
Private LastColumn&
Private SearchRange, CustomerPart As Range

Public Sub SetupNotOrdered()
    Call SetSearchHeader("Customer Part #")
    Call AddHeaders
    Call FindNotOrderedParts
    Call AddGFCPartNumbers
End Sub

Private Sub SetSearchHeader(header_text$)
    LastColumn = Sheets("Serial File").Range("IV1").End(xlToLeft).Column
    Sheets("Serial File").Select
    
    Set SearchRange = Sheets("Serial File").Range(Cells(1, 1), Cells(1, LastColumn))
    Set CustomerPart = Range(SearchRange.Find(header_text).Address)
End Sub

Private Sub AddHeaders()
    Sheets("Parts Not Ordered").Range("A1").value = "Customer Part #"
    Sheets("Parts Not Ordered").Range("B1").value = "GFC Part #"
End Sub

Private Sub FindNotOrderedParts()
    Dim i%, last_row&
    
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    For i = 2 To last_row
        'if countif(search SF column for item) = 1 and sumif( review data column for item) = 0 then
        If IsLastItemInColumn(i) = 1 Then
            If ReturnReviewSum(CustomerPart.Cells(i, 1)) = 0 Then
                Sheets("Parts Not Ordered").Range("A" & GetLastRow) = CustomerPart.Cells(i, 1)
            End If
        End If
    Next i
End Sub

Private Sub AddGFCPartNumbers()
    Dim last_row&
    
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    Sheets("Parts Not Ordered").Range("B2").FormulaR1C1 = "=VLOOKUP(RC[-1],'Serial File'!R2C5:R" & last_row & "C6,2,FALSE)"
    
    Sheets("Parts Not Ordered").Range("B2:B" & GetLastRow - 1).FillDown
    Sheets("Parts Not Ordered").Range("B2:B" & GetLastRow - 1).Copy
    Sheets("Parts Not Ordered").Range("B2").PasteSpecial Paste:=xlPasteValues
End Sub

Private Function GetLastRow%()
    GetLastRow = Sheets("Parts Not Ordered").Range("A65535").End(xlUp).Row + 1
End Function

Private Function IsLastItemInColumn%(row_num%)
    IsLastItemInColumn = Application.WorksheetFunction.CountIf(CustomerPart.EntireColumn, _
        CustomerPart.Cells(row_num, 1))
End Function

Private Function ReturnReviewSum&(part_number$)
    ReturnReviewSum = Application.WorksheetFunction.SumIf(Sheets("Review Data").Range("E:E"), _
        part_number, Sheets("Review Data").Range("X:X"))
End Function
