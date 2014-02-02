VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmColumnHide 
   Caption         =   "Choose Columns To Hide"
   ClientHeight    =   6060
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   6165
   OleObjectBlob   =   "frmColumnHide.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmColumnHide"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'
'
'Private ColHideH$()
'Private ColHideC%()
'Private ColUnHideH$()
'Private ColUnHideC%()
'
'Private Sub UserForm_Initialize()
'    lblCurrentWB02.Caption = ActiveWorkbook.Name
'    lblCurrentSheet02.Caption = activesheet.Name
'    Call Update
'End Sub
'Private Sub Update()
'    Call ClearDisplay
'    Call LoadArr
'    Call UpdateDisplay
'    ActiveWindow.ScrollColumn = 1
'End Sub
'Private Sub ClearDisplay()
'    lstUnhiddenColumns.Clear
'    lstHiddenColumns.Clear
'End Sub
'Sub LoadArr()
'    Dim TotalColumns As Byte
'    Dim x As Byte
'
'    ReDim ColHideH(0)
'    ReDim ColHideC(0)
'    ReDim ColUnHideH(0)
'    ReDim ColUnHideC(0)
'
'    TotalColumns = ActiveCell.SpecialCells(xlLastCell).Column
'
'    For x = 1 To TotalColumns
'        Select Case Columns(x).EntireColumn.Hidden
'            Case True
'                ReDim Preserve ColHideH(UBound(ColHideH) + 1)
'                ReDim Preserve ColHideC(UBound(ColHideC) + 1)
'
'                ColHideH(UBound(ColHideH)) = Range(Cells(1, x), Cells(1, x)).Value
'                ColHideC(UBound(ColHideC)) = x
'            Case False
'                ReDim Preserve ColUnHideH(UBound(ColUnHideH) + 1)
'                ReDim Preserve ColUnHideC(UBound(ColUnHideC) + 1)
'
'                ColUnHideH(UBound(ColUnHideH)) = Range(Cells(1, x), Cells(1, x)).Value
'                ColUnHideC(UBound(ColUnHideC)) = x
'        End Select
'    Next x
'End Sub
'Sub UpdateDisplay()
'    Dim HideArr As Byte
'    Dim UnhideArr As Byte
'    Dim x As Byte
'
'    HideArr = UBound(ColHideH)
'    UnhideArr = UBound(ColUnHideH)
'
'    For x = 1 To UnhideArr
'        lstUnhiddenColumns.AddItem ColUnHideH(x)
'    Next x
'
'    For x = 1 To HideArr
'        lstHiddenColumns.AddItem ColHideH(x)
'    Next x
'End Sub
'Private Sub lstHiddenColumns_Click()
'    Columns(ColHideC(lstHiddenColumns.ListIndex + 1)).EntireColumn.Hidden = False
'    Call Update
'End Sub
'Private Sub lstUnhiddenColumns_Click()
'    Columns(ColUnHideC(lstUnhiddenColumns.ListIndex + 1)).EntireColumn.Hidden = True
'    Call Update
'End Sub
'Private Sub cmdUnHideAll_Click()
'    Cells.EntireColumn.Hidden = False
'    Call Update
'End Sub
'
