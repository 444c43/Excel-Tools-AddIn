VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnique 
   Caption         =   "Unique Data Counter"
   ClientHeight    =   2400
   ClientLeft      =   45
   ClientTop       =   345
   ClientWidth     =   2460
   OleObjectBlob   =   "frmUnique.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmUnique"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private Sub UserForm_Initialize()
    Dim LastCol%
    Dim x%
    'determine last numerical column
    LastCol = ActiveCell.SpecialCells(xlLastCell).Column
    'add each column letter to combo box
    For x = 1 To LastCol
        cbxColLetters.AddItem GetColumnLetters(Columns(x).AddressLocal(ColumnAbsolute:=False))
    Next x
    'set initial value
    cbxColLetters.Text = cbxColLetters.list(0)
End Sub

Private Sub cmdCalc_Click()
    Dim UniqueItems As SheetTools
    Set UniqueItems = New SheetTools
    
    'show total unique count
    MsgBox "A total of " & UniqueItems.CountUnique(cbxColLetters.Text) & " unique items found in column " & cbxColLetters.Text & "."
End Sub

Private Function GetColumnLetters$(Value$)
    'select case if A:A or AA:AA
    Select Case Len(Value)
        Case 3
            GetColumnLetters = left(Value, 1)
        Case 5
            GetColumnLetters = left(Value, 2)
    End Select
End Function
