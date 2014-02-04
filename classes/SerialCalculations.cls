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
        
        If CheckIfScanned = False Then
            If NotScanned Then
                Sheets("Serial File").Range("W" & i).value = "Not Scanned"
            ElseIf InactiveZero Then
                Sheets("Serial File").Rows(i).Delete
                'Sheets("Serial File").Range("W" & i).value = "Inactive Zero Bins"
            ElseIf Inactive Then
                Sheets("Serial File").Range("W" & i).value = "Inactive"
            ElseIf ToReview Then
                Sheets("Serial File").Range("X" & i).value = "x" 'to review
            End If
        Else
            Sheets("Serial File").Range("W" & i).value = "Scanned"
        End If
        
        If CalculateSalesUnder15(i) Then
            Sheets("Serial File").Range("Y" & i).value = "x"
        End If
    Next i
End Sub

'Private Sub SetSearchHeader(header_text$)
'    LastColumn = Sheets("Serial File").Range("IV1").End(xlToLeft).Column
'    Sheets("Serial File").Select
'
'    Set SearchRange = Sheets("Serial File").Range(Cells(1, 1), Cells(1, LastColumn))
'    Set CustomerPart = Range(SearchRange.Find(header_text).Address)
'End Sub

Private Sub UpdateValues(i&)
    ShipQty = Sheets("Serial File").Range("G" & i).value
    BinSys = Sheets("Serial File").Range("H" & i).value
    NoCharge = Sheets("Serial File").Range("Q" & i).value
End Sub

Private Function CheckIfScanned() As Boolean
    CheckIfScanned = (ShipQty > 0)
End Function
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
    CalculateSalesUnder15 = (ShipQty > 0 And ShipQty * Sheets("Serial File").Range("S" & i).value < 15)
End Function
