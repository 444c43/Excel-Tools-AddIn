VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "PageOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub ChangeHeaders(center$, left$, right$)
    With activesheet.PageSetup
        .CenterHeader = center
        .LeftHeader = left
        .RightHeader = right
    End With
End Sub
            
Public Sub ChangeFooters(center_footer$, left_footer$, right_footer$)
    With activesheet.PageSetup
        .CenterFooter = center_footer
        .LeftFooter = left_footer
        .RightFooter = right_footer
    End With
End Sub

Public Sub ChangeLeftRightMargins(left#, right#)
    With activesheet.PageSetup
        .LeftMargin = Application.InchesToPoints(left)
        .RightMargin = Application.InchesToPoints(right)
    End With
End Sub

Public Sub ChangeTopBottomMargins(top#, bottom#)
    With activesheet.PageSetup
        .TopMargin = Application.InchesToPoints(top)
        .BottomMargin = Application.InchesToPoints(bottom)
    End With
End Sub

Public Sub ChangeHeaderFooterMargins(header#, footer#)
    With activesheet.PageSetup
        .HeaderMargin = Application.InchesToPoints(header)
        .FooterMargin = Application.InchesToPoints(footer)
    End With
End Sub

Public Sub ChangeOrientation(horizontal As Variant, vertical As Variant, orientation As Variant)
    With activesheet.PageSetup
        .CenterHorizontally = horizontal
        .CenterVertically = vertical
        .orientation = orientation
    End With
End Sub
