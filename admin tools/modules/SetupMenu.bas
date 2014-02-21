Attribute VB_Name = "SetupMenu"
Option Explicit
Private NewMenu As Menu

Public Sub CreateToolbarMenu()
    Set NewMenu = New Menu
    NewMenu.RemoveMenu
    NewMenu.AddMainMenu
    
    CreateReviewSubMenu
End Sub

Public Sub DestroyToolbarMenu()
    Set NewMenu = New Menu
    NewMenu.RemoveMenu
End Sub

Private Sub CreateReviewSubMenu()
    Call NewMenu.AddSubMenu("Macros", Array("Save Components"), Array(5593), Array("SaveComponents.EntryPoint"))
End Sub
