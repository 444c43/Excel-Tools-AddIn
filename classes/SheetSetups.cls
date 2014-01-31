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

Public Sub AdjustSheets(new_sheet_names)
    Call ApplicationProperties(False)
    Call CreateSheets(new_sheet_names, UBound(new_sheet_names))
    Call DeleteSheets(new_sheet_names)
    Call ApplicationProperties(True)
End Sub
Private Sub ApplicationProperties(bool_value As Boolean)
    Application.ScreenUpdating = bool_value
    Application.DisplayAlerts = bool_value
End Sub
Private Sub CreateSheets(new_names, sheetcount%)
    Dim i%
    For i = 0 To sheetcount
        Sheets.Add After:=Sheets(ActiveWorkbook.Sheets.count)
        activesheet.name = new_names(i)
    Next i
End Sub

Private Sub DeleteSheets(new_names)
    Dim current_sheet As Worksheet
    
    For Each current_sheet In ActiveWorkbook.Sheets
        If DoesSheetExist(current_sheet.name, new_names) = False Then
            Sheets(current_sheet.name).Delete
        End If
    Next current_sheet
End Sub

Private Function DoesSheetExist(current_name, sheet_names) As Boolean
    Dim new_sheet_name
    
    DoesSheetExist = False
    For Each new_sheet_name In sheet_names
        If CStr(new_sheet_name) = current_name Then
            DoesSheetExist = True
            Exit Function
        End If
    Next new_sheet_name
End Function
