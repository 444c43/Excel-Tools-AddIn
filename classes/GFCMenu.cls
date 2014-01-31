VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "GFCMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MainMenuBar As CommandBar
Private SubMenu As CommandBarControl
Private HelpMenu As Integer
Private SubMenuOpt As CommandBarControl

Public Sub RemoveMenu()
    'removes any pre-existing GFC menu
    On Error Resume Next
        Application.CommandBars("Worksheet Menu Bar").Controls("GFC Tools").Delete
    On Error GoTo 0
End Sub

Public Sub AddMainMenu()
    Set MainMenuBar = Application.CommandBars("Worksheet Menu Bar")
    
    'return Help menu index number
    HelpMenu = MainMenuBar.Controls("Help").Index
    
    'add GFC Tools menu before Help menu using index number
    Set SubMenu = MainMenuBar.Controls.Add(Type:=msoControlPopup, Before:=HelpMenu)
    SubMenu.Caption = "GFC Tools"
End Sub

Sub AddSubMenu(MenuTitle As String, MenuC, MenuI, MenuA)
    Dim iterate As Byte

    Set MainMenuBar = Application.CommandBars("Worksheet Menu Bar")
    Set SubMenu = MainMenuBar.Controls("GFC Tools")

    'create sub menu controls using passed options
    Set SubMenuOpt = SubMenu.Controls.Add(Type:=msoControlPopup)
    SubMenuOpt.Caption = MenuTitle
    
    For iterate = 0 To UBound(MenuC)
        With SubMenuOpt.Controls.Add(Type:=msoControlButton)
            '.Enabled = MenuAvail(iterate)
            .Caption = MenuC(iterate)
            .FaceId = MenuI(iterate)
            .OnAction = MenuA(iterate)
        End With
    Next iterate
End Sub
