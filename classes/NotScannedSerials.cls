VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "NotScannedSerials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub SetupTab()
    Dim i&, last_row&
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    Sheets("Not Scanned").Range("A1:AA1").value = Sheets("Serial File").Range("A1:AA1").value
    Sheets("Review Data").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited
    
    For i = last_row To 2 Step -1
        If BinQtyZeroAndNotScanned(i) Then
            Sheets("Serial File").Range("A" & i).EntireRow.Copy
            Sheets("Not Scanned").Range("A" & GetLastRow).PasteSpecial Paste:=xlPasteValues
        End If
    Next i
End Sub

Private Function BinQtyZeroAndNotScanned(i&) As Boolean
    BinQtyZeroAndNotScanned = (Application.WorksheetFunction.CountIf(Sheets("Review Data").Range("A:A"), Sheets("Serial File").Range("A" & i).value) = 0)
End Function

Private Function GetLastRow%()
    GetLastRow = Sheets("Not Scanned").Range("A65535").End(xlUp).Row + 1
End Function
