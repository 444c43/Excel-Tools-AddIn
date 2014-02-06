Attribute VB_Name = "MenuActions"
Option Explicit

Public Sub CountUniqueItems()
    frmUnique.Show
End Sub

Public Sub SaveSnapshot()
On Error GoTo ErrorHandler:
    Sheets("Snapshot").Copy
    ActiveWorkbook.Application.Dialogs(xlDialogSaveAs).Show
    ActiveWorkbook.Close
    Exit Sub
ErrorHandler:
    MsgBox ("No Snapshot Available. Check Sheet Names!")
End Sub
