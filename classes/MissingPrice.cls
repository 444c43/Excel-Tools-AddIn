VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MissingPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private FromRange, SearchRange, GFCPart, PriceCompare As Range

Public Sub SetupMissingPcPrice()
    Call SetSearchHeader("GFC Part #")
    Call AddHeaders
    Call FindMissingPcPrice
End Sub

Private Sub SetSearchHeader(header_text$)
    Dim LastColumn%
    LastColumn = Sheets("Serial File").Range("IV1").End(xlToLeft).Column
    
    Sheets("Serial File").Select
    Set FromRange = Sheets("Serial File").Range(Cells(1, 1), Cells(1, LastColumn))
    Set GFCPart = Sheets("Serial File").Range(FromRange.Find(header_text).Address)
    
    Sheets("Price List").Select
    Set SearchRange = Sheets("Price List").Range(Cells(1, 1), Cells(1, LastColumn))
    Set PriceCompare = Sheets("Price List").Range(SearchRange.Find(header_text).Address)
End Sub

Private Sub AddHeaders()
    Sheets("Missing Pc Price").Range("A1").Value = "GFC Part #"
    Sheets("Missing Pc Price").Range("B1").Value = "Customer Part #"
End Sub

Private Sub FindMissingPcPrice()
    Dim i%, last_row&, row_num%
    last_row = Sheets("Review Data").Range("A65535").End(xlUp).Row
    
    For i = 2 To last_row
        If IsLastPart(i) = 1 Then
            If EvaluatePcPrice(i) = 0 Then
                row_num = GetLastRow + 1
                Sheets("Missing Pc Price").Range("A" & row_num).Value = GFCPart.Cells(i, 1)
                Sheets("Missing Pc Price").Range("B" & row_num).Value = GFCPart.Cells(i, 0)
            End If
        End If
    Next i
End Sub

Private Function IsLastPart%(i%)
    IsLastPart = Application.WorksheetFunction.CountIf(GFCPart.EntireColumn, GFCPart.Cells(i, 1))
End Function
Private Function EvaluatePcPrice%(i%)
    PriceCompare.EntireColumn.Select
    EvaluatePcPrice = Application.WorksheetFunction.CountIf(PriceCompare.EntireColumn, GFCPart.Cells(i, 1))
End Function
Private Function GetLastRow%()
    GetLastRow = Sheets("Missing Pc Price").Range("A65535").End(xlUp).Row
End Function

