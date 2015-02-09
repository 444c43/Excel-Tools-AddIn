Attribute VB_Name = "SetupMenu"
Option Explicit
Private NewMenu As GFCMenu

Public Sub CreateToolbarMenu()
    Set NewMenu = New GFCMenu
    NewMenu.RemoveMenu
    NewMenu.AddMainMenu
    
    CreateReviewSubMenu
    CreateSnapshotSubMenu
    CreateUniqueToolsSubMenu
    CreateSFEitSubMenu
    CreateLblPrintSubMenu
    CreateBarCoderSubMenu
End Sub

Public Sub DestroyToolbarMenu()
    Set NewMenu = New GFCMenu
    NewMenu.RemoveMenu
End Sub

Private Sub CreateReviewSubMenu()
    Call NewMenu.AddSubMenu("Review", _
        Array("Setup Review", "Zero Backfill", "Run Review", "Version 8.0"), _
        Array(5593, 398, 3524, 3998), _
        Array("RunReview.SetupSheets", "ZeroBackfill.BackfillCells", "RunReview.EntryPoint", ""))
End Sub
Private Sub CreateSnapshotSubMenu()
    Call NewMenu.AddSubMenu("Snapshot", _
        Array("Run Snapshot", "Publish Snapshot", "Import Snapshot", "Export Snapshot", "Version 1.0"), _
        Array(3524, 284, 106, 1679, 3998), _
        Array("Snapshot.Run", "MenuActions.SaveSnapshot", "MenuActions.ImportSnapshot", "MenuActions.ExportSnapshot", ""))
End Sub
Private Sub CreateSFEitSubMenu()
    Call NewMenu.AddSubMenu("SF Editing", _
        Array("SF BUILD", "SF UPDATE"), _
        Array(1399, 1396), _
        Array("AS400Labels.SFBUILD", "AS400Labels.SFUPDATE"))
End Sub
Private Sub CreateLblPrintSubMenu()
    Call NewMenu.AddSubMenu("AS400 Lbl Printing", _
        Array("3RDPARTY", "HARRYO2Z", "ONELINE", "STAND1X3"), _
        Array(509, 509, 509, 509), _
        Array("AS400Labels.THIRDPARTY", "AS400Labels.HARRYO2Z", "AS400Labels.ONELINE", "AS400Labels.STAND1X3"))
End Sub
Private Sub CreateUniqueToolsSubMenu()
    Call NewMenu.AddSubMenu("Unique Tools", _
        Array("Rename AS400 Headers", "Hide/Unhide Columns", "Unique Items In Column", "Pane Freeze/Unfreeze @ A2"), _
        Array(1549, 9, 4153, 1742), _
        Array("MenuActions.ConvertHeaders", "MenuActions.HideUnhideColumns", "MenuActions.CountUniqueItems", "MenuActions.PaneFreeze"))
End Sub
Private Sub CreateBarCoderSubMenu()
    Call NewMenu.AddSubMenu("Bar Coder", _
        Array("Setup Bar Coder", "Run Bar Coder", "Version 1.0"), _
        Array(627, 19, 498), _
        Array("", "", ""))
End Sub
