VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetSetups"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub AdjustSheets(Names)
    Application.ScreenUpdating = False
    Call AddAndRemove(UBound(Names) + 1)
    Call Rename(Names)
    Sheets(1).Select
    Cells.Delete
    Application.ScreenUpdating = True
End Sub

Private Sub AddAndRemove(SheetCountMax)
' adds or deletes sheets based on sent data
    'delete all but one sheet
    Do While ActiveWorkbook.Sheets.count > 1
        Application.DisplayAlerts = False
        ActiveWindow.SelectedSheets.Delete
        Application.DisplayAlerts = True
    Loop
    'add needed sheets
    Do While ActiveWorkbook.Sheets.count < SheetCountMax
        Sheets.Add After:=Sheets(ActiveWorkbook.Sheets.count)
    Loop
End Sub

Private Sub Rename(SheetNamesArray)
    Dim x As Byte
    'rename all sheets accordingly
    For x = 1 To ActiveWorkbook.Sheets.count
        Sheets(x).name = SheetNamesArray(x - 1)
    Next x
End Sub

