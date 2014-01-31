Attribute VB_Name = "SheetTools"
Option Explicit
Sub ShowUnique()
    frmUnique.Show
End Sub
Function UniqueCount(ShtVal As String, ColLet As String, endrge As Integer)
    Dim ShtName As Object
    Dim PartTotal As Integer
    Dim Current As Integer
    Dim x As Integer
    
    PartTotal = 0
    Set ShtName = Sheets(ShtVal)
    For x = 2 To endrge
        Current = Application.WorksheetFunction.CountIf(ShtName.Range(ColLet & x & ":" & ColLet & endrge), ShtName.Range(ColLet & x))
        If Current = 1 Then
            PartTotal = PartTotal + 1
        End If
    Next x
    UniqueCount = PartTotal
End Function
Sub PaneFreeze()
    Range("A2").Select
    If ActiveWindow.FreezePanes = True Then
        ActiveWindow.FreezePanes = False
    Else
        ActiveWindow.FreezePanes = True
    End If
End Sub
