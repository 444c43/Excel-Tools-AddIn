Attribute VB_Name = "MenuSetup"
Option Explicit
Private NewMenu As GFCMenu
Private Titles()
Private Icons()
Private Actions()

Public Sub CreateToolbarMenu()
    Set NewMenu = New GFCMenu
    NewMenu.RemoveMenu
    NewMenu.AddMainMenu
    
    CreateReviewSubMenu
    CreateSnapshotSubMenu
    CreateUniqueToolsSubMenu
    CreateLblPrintSubMenu
'    CreateBarCoderSubMenu
End Sub

Public Sub DestroyToolbarMenu()
    Set NewMenu = New GFCMenu
    NewMenu.RemoveMenu
End Sub

Private Sub CreateReviewSubMenu()
    Titles = Array("Setup Review", "Run Review", "Version 8.0")
    Icons = Array(5593, 3524, 3998)
    Actions = Array("Review.SetupReviewSheets", "Review.Run", "")
    
    Call NewMenu.AddSubMenu("Review", Titles, Icons, Actions)
End Sub

Private Sub CreateSnapshotSubMenu()
    Titles = Array("Run Snapshot", "Publish Snapshot", "Import Snapshot", "Export Snapshot", "Version 1.0")
    Icons = Array(3524, 284, 106, 1679, 3998)
    Actions = Array("ReviewSnapshot.Run", "MenuSetup.SaveSnapshot", "SnapshotImportExport.Import", "SnapshotImportExport.Export", "")
    
    Call NewMenu.AddSubMenu("Snapshot", Titles, Icons, Actions)
End Sub
Private Sub CreateLblPrintSubMenu()
    Titles = Array("3RDPARTY", "SF BUILD", "HARRYO2Z", "SF UPDATE", "STAND1X3", "ONELINE")
    Icons = Array(509, 509, 509, 509, 509, 509)
    Actions = Array("AS400Labels.THIRDPARTY", "AS400Labels.SFBUILD", "AS400Labels.HARRYO2Z", "SFUPDATE", "STAND1X3", "ONELINE")
    
    Call NewMenu.AddSubMenu("AS400 Labels", Titles, Icons, Actions)
End Sub

Private Sub CreateUniqueToolsSubMenu()
    Titles = Array("Rename AS400 Headers", "Hide/Unhide Columns", "Unique Items In Column", "Pane Freeze/Unfreeze @ A2")
    Icons = Array(1549, 9, 4153, 1742)
    Actions = Array("MenuSetup.RenameHeaders", "ColumnHide.Run", "SheetTools.ShowUnique", "SheetTools.PaneFreeze")
    
    Call NewMenu.AddSubMenu("Unique Tools", Titles, Icons, Actions)
End Sub

'Private Sub CreateBarCoderSubMenu()
'    Titles = Array("Setup Bar Coder", "Run Bar Coder", "Version 1.0")
'    Icons = Array(627, 19, 498)
'    Actions = Array("Setup.Review", "Review.Run", "")
'
'    Call NewMenu.AddSubMenu("Bar Coder", Titles, Icons, Actions)
'End Sub
'
Sub RenameHeaders()
    TextFormatting.HeaderCorrect
    SheetFormatting.AllCellsFit
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
