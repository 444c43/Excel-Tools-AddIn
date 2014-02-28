VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SerialCalculations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private ShipQty#, BinSys%, NoCharge$
    
Sub CalculateSerials()
    Dim i&, last_row&
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    For i = last_row To 2 Step -1
        Call UpdateValues(i)
        
        If (ShipQty > 0) Then
            Call AddStatusValue(i, "Scanned")
        Else
            If NotScanned Then
                Call AddStatusValue(i, "Not Scanned")
                Call CopyNotScanned(i)
            ElseIf InactiveZero Then
                Sheets("Serial File").Rows(i).Delete
            ElseIf Inactive Then
                Call AddStatusValue(i, "Inactive")
            ElseIf ToReview Then
                Sheets("Serial File").Range("X" & i).Value = "x"
            End If
        End If
        
        If CalculateSalesUnder15(i) Then
            Sheets("Serial File").Range("Y" & i).Value = "x"
        End If
    Next i
End Sub
Private Sub AddStatusValue(i&, status$)
    Sheets("Serial File").Range("W" & i).Value = status
End Sub
Private Sub CopyNotScanned(i&)
    Sheets("Serial File").Range("A" & i).EntireRow.Copy
    Sheets("Not Scanned").Range("A" & GetLastRow("Not Scanned")).PasteSpecial Paste:=xlPasteValues
End Sub
Private Function GetLastRow%(current_sheet$)
    GetLastRow = Sheets(current_sheet).Range("A65535").End(xlUp).Row + 1
End Function

Sub MoveSerialStatus()
    Sheets("Serial File").Columns("W:W").Cut
    Sheets("Serial File").Columns("A:A").Insert Shift:=xlToRight
End Sub

Private Sub UpdateValues(i&)
    ShipQty = Sheets("Serial File").Range("G" & i).Value
    BinSys = Sheets("Serial File").Range("H" & i).Value
    NoCharge = Sheets("Serial File").Range("Q" & i).Value
End Sub

Private Function NotScanned() As Boolean
    NotScanned = (ShipQty = 0 And NoCharge <> "I" And BinSys > 0)
End Function
Private Function InactiveZero() As Boolean
    InactiveZero = (ShipQty = 0 And BinSys = 0)
End Function
Private Function Inactive() As Boolean
    Inactive = (ShipQty = 0 And NoCharge = "I" And BinSys > 0)
End Function
Private Function ToReview() As Boolean
    ToReview = (ShipQty > 0 And NoCharge = "I" And BinSys > 0 Or ShipQty > 0 And NoCharge <> "I" And BinSys = 0 Or _
        ShipQty > 0 And NoCharge = "I" And BinSys = 0 Or ShipQty = 0 And NoCharge <> "I" And BinSys = 0)
End Function
Private Function CalculateSalesUnder15(i&)
    CalculateSalesUnder15 = (ShipQty > 0 And ShipQty * Sheets("Serial File").Range("S" & i).Value < 15)
End Function
