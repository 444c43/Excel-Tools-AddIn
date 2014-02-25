Attribute VB_Name = "MenuActions"
Option Explicit
Private SnapshotImportExport As ImpExpSnapshot

Public Sub CountUniqueItems()
    frmUnique.Show
End Sub

Public Sub ConvertHeaders()
    Dim SheetHeaders As AS400
    
    Set SheetHeaders = New AS400
    
    Call SheetHeaders.ConvertHeaders(activesheet.name)
End Sub

Public Sub HideUnhideColumns()
    frmColumnHide.Show
End Sub

Sub PaneFreeze()
    Range("A2").Select
    If ActiveWindow.FreezePanes = True Then
        ActiveWindow.FreezePanes = False
    Else
        ActiveWindow.FreezePanes = True
    End If
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
