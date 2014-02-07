Attribute VB_Name = "MenuActions"
Option Explicit
Private SnapshotImportExport As ImpExpSnapshot

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

Public Sub ImportSnapshot()
    Set SnapshotImportExport = New ImpExpSnapshot
    SnapshotImportExport.Import
End Sub

Public Sub ExportSnapshot()
    Set SnapshotImportExport = New ImpExpSnapshot
    SnapshotImportExport.Export
End Sub
