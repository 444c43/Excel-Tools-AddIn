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

Private Function IsSerialNotScanned(i%) As Boolean
    IsSerialNotScanned = (Application.WorksheetFunction.CountIf(Sheets("Review Data").Range("A:A"), Sheets("Serial File").Range("A" & i).Value) = 0)
End Function

'INACTIVE SERIALS
Public Sub CutDeletedCopyInactive()
    For i = last_row To 2 Step -1
        Call EvaluateSerialStatus(i, GetLastRow("Inactive Serials"))
    Next i
    Call AddPcPrice
    Call FillDownCopyAndPaste
    Call RemoveNA
End Sub

Private Sub EvaluateSerialStatus(i%, inactive_last_row%)
    Select Case Sheets("Serial File").Range("W" & i).Value
        Case "D"
            Sheets("Serial File").Rows(i).Cut
            Sheets("Inactive Serials").Range("A" & inactive_last_row).Select
            activesheet.Paste
        Case "I"
            Sheets("Serial File").Rows(i).Cut
            Sheets("Inactive Serials").Range("A" & inactive_last_row).Select
            activesheet.Paste
    End Select
    If Sheets("Serial File").Range("Q" & i).Value = "I" Then
        Sheets("Serial File").Rows(i).Cut
        Sheets("Inactive Serials").Range("A" & inactive_last_row).Select
        activesheet.Paste
    End If
End Sub

Private Sub AddPcPrice()
    With Sheets("Inactive Serials")
        .Range("AB1").Value = "Pc Price"
        .Range("AB2").FormulaR1C1 = "=VLOOKUP(RC[-22],'Price List'!R2C3:R" & GetLastRow("Price List") & "C6,4,False)"
        .Range("AC1").Value = "Loop Value"
        .Range("AC2").FormulaR1C1 = "=SUM(RC[-1]*RC[-14]*RC[-22])"
    End With
End Sub

Private Sub FillDownCopyAndPaste()
    With Sheets("Inactive Serials")
        .Range("AB2:AC" & GetLastRow("Inactive Serials")).FillDown
        .Columns("AB:AC").Copy
        .Range("AB1").PasteSpecial Paste:=xlPasteValues
    End With
End Sub

Private Sub RemoveNA()
    Sheets("Inactive Serials").Columns("AB:AC").Replace What:="#N/A", Replacement:="", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
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
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
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
Public Sub SetupPartsNotOrdered(not_ordered As Collection)
    Dim last_row&
    
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    Call AddHeaders("Parts Not Ordered", "Customer Part #", "GFC Part #")
    Call AddCustomerPartNumbers(not_ordered)
    Call AddGFCPartNumbers
    Call TextToColumns
    Call CheckForErrors
    Call CopyPaste
End Sub
Private Sub AddCustomerPartNumbers(not_ordered As Collection)
    Dim x%
    For x = 1 To not_ordered.count
        If (Sheets.Application.WorksheetFunction.CountIf(Sheets("Review Data").Range("E:E"), not_ordered(x)) = 0) Then
            Sheets("Parts Not Ordered").Range("A" & LastPNOCell()).Value = not_ordered(x)
        End If
    Next x
End Sub
Private Function LastPNOCell() As Integer
    LastPNOCell = Sheets("Parts Not Ordered").Range("A65535").End(xlUp).Row + 1
End Function
Private Sub AddGFCPartNumbers()
    Sheets("Parts Not Ordered").Range("B2").FormulaR1C1 = "=VLOOKUP(RC[-1],'Serial File'!R2C4:R" & last_row & "C5,2,FALSE)"
    Sheets("Parts Not Ordered").Range("B2:B" & GetLastRow("Parts Not Ordered") - 1).FillDown
End Sub
Private Sub TextToColumns()
    Sheets("Inactive Serials").Columns("E:E").TextToColumns Destination:=Range("E1"), DataType:=xlDelimited, _
        TextQualifier:=xlDoubleQuote, ConsecutiveDelimiter:=False, Tab:=False, _
        Semicolon:=False, Comma:=False, Space:=False, Other:=False, FieldInfo _
        :=Array(1, 1), TrailingMinusNumbers:=True
End Sub
Private Sub CheckForErrors()
    Dim x%
    For x = 2 To GetLastRow("Parts Not Ordered") - 1
        If IsError(Sheets("Parts Not Ordered").Range("B" & x).Value) Then
            Sheets("Parts Not Ordered").Range("B" & x).FormulaR1C1 = _
                "=VLOOKUP(RC[-1],'Inactive Serials'!R2C5:R" & GetLastRow("Inactive Serials") & "C6,2,FALSE)"
        End If
    Next x
End Sub
Private Sub CopyPaste()
    Sheets("Parts Not Ordered").Range("B2:B" & GetLastRow("Parts Not Ordered") - 1).Copy
    Sheets("Parts Not Ordered").Range("B2").PasteSpecial Paste:=xlPasteValues
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
