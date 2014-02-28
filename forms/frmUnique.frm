VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmUnique 
   Caption         =   "Unique Data Counter"
   ClientHeight    =   2475
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
Public TotalItems As UniqueContent

Private Sub UserForm_Initialize()
    Set TotalItems = New UniqueContent
    Call AddHeadersToComboBox
    cbxColLetters.Text = cbxColLetters.list(0)
End Sub

Private Sub cmdCalc_Click()
    Call TotalItems.Initialize(activesheet.name, cbxColLetters.Text)
    Call TotalItems.GetUniqueList
    Call DisplayResults
End Sub

Private Sub AddHeadersToComboBox()
    Dim i As Byte
    For i = 1 To GetLastColumn()
        cbxColLetters.AddItem GetHeaderName(i)
    Next i
End Sub

Private Function GetHeaderName$(i As Byte)
    GetHeaderName = Range(Cells(1, i), Cells(1, i)).Value
End Function

Private Function GetLastColumn() As Byte
    GetLastColumn = ActiveCell.SpecialCells(xlLastCell).Column
End Function

Private Sub DisplayResults()
    MsgBox "Column with header: " & cbxColLetters.Text & " has " & TotalItems.UniqueItems.count & " unique item(s)"
End Sub

