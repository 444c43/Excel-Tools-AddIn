VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReviewSheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private i%, last_row&
Private LastColumn&
Private SearchRange, CustomerPart As Range
Private FromRange, GFCPart, PriceCompare As Range

Public Sub AddNewSheet(new_sheet_name$, after_sheet_name$)
    Sheets.Add After:=Sheets(after_sheet_name)
    activesheet.name = new_sheet_name
End Sub

Public Sub InstantiateVariables()
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
End Sub

'SERIALS NOT SCANNED
Public Sub CopyPasteNotScanned()
    Sheets("Review Data").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited
    
    For i = last_row To 2 Step -1
        If IsSerialNotScanned(i) Then
            Sheets("Serial File").Range("A" & i).EntireRow.Copy
            Sheets("Not Scanned").Range("A" & GetLastRow("Not Scanned")).PasteSpecial Paste:=xlPasteValues
        End If
    Next i
End Sub
Private Function IsSerialNotScanned(i%) As Boolean
    IsSerialNotScanned = (Application.WorksheetFunction.CountIf(Sheets("Review Data").Range("A:A"), Sheets("Serial File").Range("A" & i).Value) = 0)
End Function

'INACTIVE SERIALS
Public Sub CutDeletedCopyInactive()
    For i = last_row To 2 Step -1
        Call EvaluateSerialStatus(i, GetLastRow("Inactive Serials"))
    Next i
End Sub

Private Sub EvaluateSerialStatus(i%, inactive_last_row%)
    Select Case Sheets("Serial File").Range("W" & i).Value
    Case "D"
        Sheets("Serial File").Rows(i).Cut
        Sheets("Inactive Serials").Range("A" & inactive_last_row).Select
        activesheet.Paste
    Case "I"
        Sheets("Serial File").Rows(i).Copy
        Sheets("Inactive Serials").Range("A" & inactive_last_row).Select
        activesheet.Paste
    End Select
End Sub

'MISSING PRICE
Public Sub SetupMissingPcPrice()
    Call SetMissingSearchHeader("GFC Part #")
    Call AddHeaders("Missing Pc Price", "GFC Part #", "Customer Part #")
    Call FindMissingPcPrice
End Sub
Private Sub SetMissingSearchHeader(header_text$)
    Dim LastColumn%
    LastColumn = Sheets("Serial File").Range("IV1").End(xlToLeft).Column
    
    Sheets("Serial File").Select
    Set FromRange = Sheets("Serial File").Range(Cells(1, 1), Cells(1, LastColumn))
    Set GFCPart = Sheets("Serial File").Range(FromRange.Find(header_text).Address)
    
    Sheets("Price List").Select
    Set SearchRange = Sheets("Price List").Range(Cells(1, 1), Cells(1, LastColumn))
    Set PriceCompare = Sheets("Price List").Range(SearchRange.Find(header_text).Address)
End Sub
Private Sub FindMissingPcPrice()
    Dim i%, last_row&, row_num%
    last_row = Sheets("Review Data").Range("A65535").End(xlUp).Row
    
    For i = 2 To last_row
        If IsLastPart(i) = 1 Then
            If EvaluatePcPrice(i) = 0 Then
                row_num = GetLastRow("Missing Pc Price")
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


'PARTS NOT ORDERED
Public Sub SetupNotOrdered()
    Call SetSearchHeader("Customer Part #")
    Call AddHeaders("Parts Not Ordered", "Customer Part #", "GFC Part #")
    Call FindNotOrderedParts
    Call AddGFCPartNumbers
End Sub
Private Sub SetSearchHeader(header_text$)
    LastColumn = Sheets("Serial File").Range("IV1").End(xlToLeft).Column
    Sheets("Serial File").Select
    
    Set SearchRange = Sheets("Serial File").Range(Cells(1, 1), Cells(1, LastColumn))
    Set CustomerPart = Range(SearchRange.Find(header_text).Address)
End Sub
Private Sub AddHeaders(sheet_name$, CellA1$, CellB1$)
    Sheets(sheet_name).Range("A1").Value = CellA1$
    Sheets(sheet_name).Range("B1").Value = CellB1$
End Sub
Private Sub FindNotOrderedParts()
    Dim i%, last_row&
    
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    For i = 2 To last_row
        'if countif(search SF column for item) = 1 and sumif( review data column for item) = 0 then
        If IsLastItemInColumn(i) = 1 Then
            If ReturnReviewSum(CustomerPart.Cells(i, 1)) = 0 Then
                Sheets("Parts Not Ordered").Range("A" & GetLastRow("Parts Not Ordered")) = CustomerPart.Cells(i, 1)
            End If
        End If
    Next i
End Sub
Private Sub AddGFCPartNumbers()
    Dim last_row&
    
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    Sheets("Parts Not Ordered").Range("B2").FormulaR1C1 = "=VLOOKUP(RC[-1],'Serial File'!R2C5:R" & last_row & "C6,2,FALSE)"
    
    Sheets("Parts Not Ordered").Range("B2:B" & GetLastRow("Parts Not Ordered") - 1).FillDown
    Sheets("Parts Not Ordered").Range("B2:B" & GetLastRow("Parts Not Ordered") - 1).Copy
    Sheets("Parts Not Ordered").Range("B2").PasteSpecial Paste:=xlPasteValues
End Sub
Private Function IsLastItemInColumn%(row_num%)
    IsLastItemInColumn = Application.WorksheetFunction.CountIf(CustomerPart.EntireColumn, _
        CustomerPart.Cells(row_num, 1))
End Function
Private Function ReturnReviewSum&(part_number$)
    ReturnReviewSum = Application.WorksheetFunction.SumIf(Sheets("Review Data").Range("E:E"), _
        part_number, Sheets("Review Data").Range("X:X"))
End Function

'SHARED PUBLIC SUBS
Public Sub CopyHeaders(from_sheet_name$, to_sheet_name$)
    Sheets(to_sheet_name).Range("A1:AA1").Value = Sheets(from_sheet_name).Range("A1:AA1").Value
End Sub

'SHARED PRIVATE SUBS
Private Function GetLastRow%(current_sheet$)
    GetLastRow = Sheets(current_sheet).Range("A65535").End(xlUp).Row + 1
End Function
