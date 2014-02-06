Attribute VB_Name = "SheetTools"
Option Explicit
Sub ShowUnique()
    frmUnique.Show
End Sub
Function UniqueCount(ShtVal$, ColLet$, endrge%)
    Dim ShtName As Object
    Dim PartTotal%
    Dim Current%
    Dim x%
    
    PartTotal = 0
    Set ShtName = sheets(ShtVal)
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
