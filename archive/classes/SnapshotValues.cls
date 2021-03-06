VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SnapshotValues"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'range, total columns and rows
Private LastColumn%, LastRow%
Private SearchRange As Range
Private WSFunction As Object
'serial values
Private scanned_total%, not_scanned_total%, inactive_total%, missing_price%
'   serial value headers
Private MissingPcPrice As Range
Private SerialStatus As Range
'part values
Private ordered_parts%, not_ordered_parts%
'   part value headers
Private CustomerPart As Range
Private ShipQty As Range
'loop values
Private sales_value@, not_scanned_value@, inactive_value@
Private BinQty, BinSys, PcPrice, NoCharge As Range
Private SheetName$

Public Sub SetProperties(SheetRef$)
    sheets(SheetRef).Select
    LastRow = ActiveCell.SpecialCells(xlLastCell).Row
    LastColumn = Range("IV1").End(xlToLeft).Column
    
    Set SearchRange = Range(Cells(1, 1), Cells(1, LastColumn))
    Set WSFunction = Application.WorksheetFunction
    
    Call SetSerialValues
    Call SetPartValues
    Call SetLoopValues
    sheets("Snapshot").Select
End Sub

Private Sub SetSerialValues()
    'set range
    Set SerialStatus = Range(SearchRange.Find("Serial Status").Address)
    Set MissingPcPrice = Range(SearchRange.Find("Pc Price").Address)
    
    ' set values
    scanned_total = WSFunction.CountIf(SerialStatus.EntireColumn, "Scanned")
    not_scanned_total = WSFunction.CountIf(SerialStatus.EntireColumn, "Not Scanned")
    inactive_total = WSFunction.CountIf(SerialStatus.EntireColumn, "Inactive")
    missing_price = WSFunction.CountIf(MissingPcPrice.EntireColumn, 0)
End Sub

Private Sub SetPartValues()
    Dim i%
    
    Set CustomerPart = Range(SearchRange.Find("Customer Part #").Address)
    Set ShipQty = Range(SearchRange.Find("Ship Qty").Address)
    
    ordered_parts = 0
    not_ordered_parts = 0
   
    For i = 2 To LastRow
        If CheckForLastListItem(i) = 1 Then
            If HasShipQty(i) Then
                ordered_parts = ordered_parts + 1
            Else
                not_ordered_parts = not_ordered_parts + 1
            End If
        End If
    Next i
End Sub
Private Function CheckForLastListItem(i%)
    CheckForLastListItem = WSFunction.CountIf(CustomerPart.Range(Cells(i, 1), Cells(LastRow, 1)), CustomerPart.Cells(i, 1))
End Function

Private Function HasShipQty(i%) As Boolean
    If WSFunction.SumIf(CustomerPart.EntireColumn, CustomerPart.Cells(i, 1), ShipQty.EntireColumn) = 0 Then
        HasShipQty = False
        Exit Function
    End If
    HasShipQty = True
End Function

Private Sub SetLoopValues()
Dim i%

Set BinQty = Range(SearchRange.Find("Bin Qty").Address)
Set BinSys = Range(SearchRange.Find("Bin Sys").Address)
Set PcPrice = Range(SearchRange.Find("Pc Price").Address)
Set NoCharge = Range(SearchRange.Find("No Charge").Address)

For i = 2 To LastRow
    If ShipQty.Cells(i, 1) > 0 Then
        sales_value = sales_value + GetSalesValue(i)
    ElseIf NoCharge.Cells(i, 1) = "I" Then
        inactive_value = inactive_value + GetBinSysValue(i)
    Else
        not_scanned_value = not_scanned_value + GetBinSysValue(i)
    End If
Next i

End Sub

Private Function GetSalesValue@(i%)
    GetSalesValue = ShipQty.Cells(i, 1) * PcPrice.Cells(i, 1)
End Function

Private Function GetBinSysValue@(i%)
    GetBinSysValue = BinQty.Cells(i, 1) * BinSys.Cells(i, 1) * PcPrice.Cells(i, 1)
End Function

'PRIVATELY SET PROPERTIES
Public Property Get Scanned%()
    Scanned = scanned_total
End Property

Public Property Get NotScanned%()
    NotScanned = not_scanned_total
End Property

Public Property Get Inactive%()
    Inactive = inactive_total
End Property

Public Property Get Missing%()
    Missing = missing_price
End Property

Public Property Get OrderedParts%()
    OrderedParts = ordered_parts
End Property

Public Property Get NotOrderedParts%()
    NotOrderedParts = not_ordered_parts
End Property

Public Property Get SalesValue@()
    SalesValue = sales_value
End Property

Public Property Get NotScannedValue@()
    NotScannedValue = not_scanned_value
End Property

Public Property Get InactiveValue@()
    InactiveValue = inactive_value
End Property
