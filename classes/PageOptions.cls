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

Public Sub ChangeHeaders(center As String, left As String, right As String)
    With activesheet.PageSetup
        'headers
        .CenterHeader = ""
        .LeftHeader = ""
        .RightHeader = ""
    End With
End Sub
            
Public Sub ChangeFooters(center As String, left As String, right As String)
    With activesheet.PageSetup
        .CenterFooter = ""
        .LeftFooter = ""
        .RightFooter = ""
    End With
End Sub

Public Sub ChangeMargins(left As Double, right As Double, top As Double, bottom As Double)
    '.75, .75, 1, 1
    With activesheet.PageSetup
        .LeftMargin = Application.InchesToPoints(left)
        .RightMargin = Application.InchesToPoints(right)
        .TopMargin = Application.InchesToPoints(top)
        .BottomMargin = Application.InchesToPoints(bottom)
    End With
End Sub

Public Sub orientation(horizontal As Variant, vertical As Variant, orientation As Variant)
    With activesheet.PageSetup
        .CenterHorizontally = horizontal
        .CenterVertically = vertical
        .orientation = orientation
    End With
End Sub
