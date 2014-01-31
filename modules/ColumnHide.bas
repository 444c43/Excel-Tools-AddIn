Attribute VB_Name = "ColumnHide"
Option Explicit

Private ColHideH() As String
Private ColHideC() As Integer
Private ColUnHideH() As String
Private ColUnHideC() As Integer

Sub Run()
    frmColumnHide.Show
    Call LoadArr
End Sub
Sub LoadArr()
    Dim TotalColumns As Byte
    Dim x As Byte
    
    ReDim ColHideH(0)
    ReDim ColHideC(0)
    ReDim ColUnHideH(0)
    ReDim ColUnHideC(0)
    
    TotalColumns = ActiveCell.SpecialCells(xlLastCell).Column
    On Error GoTo ErrorControl
    For x = 1 To TotalColumns
        Select Case Columns(x).EntireColumn.Hidden
            Case True
                ReDim Preserve ColHideH(UBound(ColHideH) + 1)
                ReDim Preserve ColHideC(UBound(ColHideC) + 1)

                ColHideH(UBound(ColHideH)) = Range(Cells(1, x), Cells(1, x)).Value
                ColHideC(UBound(ColHideC)) = x
            Case False
                ReDim Preserve ColUnHideH(UBound(ColUnHideH) + 1)
                ReDim Preserve ColUnHideC(UBound(ColUnHideC) + 1)

                ColUnHideH(UBound(ColUnHideH)) = Range(Cells(1, x), Cells(1, x)).Value
                ColUnHideC(UBound(ColUnHideC)) = x
        End Select
    Next x
    Call UpdateDisplay
ErrorControl:
End Sub
Sub UpdateDisplay()
Dim HideArr As Byte
Dim UnhideArr As Byte
Dim x As Byte

HideArr = UBound(ColHideH)
HideArr = UBound(ColUnHideH)

For x = 1 To HideArr
    frmColumnHide.lstHiddenColumns.AddItem ColHideH(x)
Next x

For x = 1 To UnhideArr
    frmColumnHide.lstUnhiddenColumns.AddItem ColUnHideH(x)
Next x



End Sub
