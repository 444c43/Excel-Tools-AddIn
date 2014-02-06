VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "InactiveSerials"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CutDeletedCopyInactive()
    Dim i%, last_row&
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    Sheets("Inactive Serials").Range("A1:AA1").Value = Sheets("Serial File").Range("A1:AA1").Value
    
    For i = last_row To 2 Step -1
        Call EvaluateSerialStatus(i, GetLastRow)
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

Private Function GetLastRow%()
    GetLastRow = Sheets("Inactive Serials").Range("A65535").End(xlUp).Row + 1
End Function
