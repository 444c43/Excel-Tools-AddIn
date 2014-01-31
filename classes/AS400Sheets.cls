VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AS400Sheets"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private AS400Labels As SheetSetups

Public Sub SetupAS400Sheets(sheet_names, sheet_headers)
    Set AS400Labels = New SheetSetups
    
    Call AS400Labels.AdjustSheets(sheet_names)
    Call AddHeaderValues(sheet_headers)
    Call FormatHeaders(sheet_headers)
    Call FormatSheet.ResizeAllCells
End Sub

Private Sub AddHeaderValues(headers)
    Dim x As Byte
    For x = 0 To UBound(headers)
        Range(Cells(1, x + 1), Cells(1, x + 1)).Value = headers(x)
    Next x
End Sub

Private Sub FormatHeaders(headers)
    Range("A2").Formula = "=SUM(65535-COUNTBLANK(B:B))+1"
    Range("A1:A2").Font.Bold = True
    Range("A1:A2").Font.ColorIndex = 3
    Range(Cells(1, 2), Cells(1, UBound(headers) + 1)).Font.Bold = True
End Sub
