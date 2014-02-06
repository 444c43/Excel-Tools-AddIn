VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "TextFormat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub NmrFmt(Format As String, AlignDirection As Variant)
    'sets number format
    Selection.NumberFormat = Format
    Selection.HorizontalAlignment = xlRight
End Sub
Public Sub TextInCell(rge As String, FontS As Integer, FontC As Integer, FontB As Boolean, FontV As String)
    Range(rge).Select
    With Selection
        .Font.Size = FontS
        .Font.ColorIndex = FontC
        .Font.Bold = FontB
        .Value = FontV
    End With
End Sub
Public Sub ObjectFont(DocObject As Object, FontName As String, FontStyle As String, FontSize As Integer, FontC As Variant)
    With DocObject
        .name = FontName
        .FontStyle = FontStyle
        .Size = FontSize
        .ColorIndex = FontC
    End With
End Sub

