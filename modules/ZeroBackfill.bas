Attribute VB_Name = "ZeroBackfill"
Option Explicit

Public Sub BackfillCells()
    Call IterateSheets
    Sheets("Serial File").Select
End Sub

Private Sub IterateSheets()
    Call ReplaceBlanks("Serial File", "G:G, O:O")
    Call ReplaceBlanks("Review Data", "G:G, O:O, X:X, Y:Y")
    Call ReplaceBlanks("Price List", "F:F")
End Sub

Private Sub ReplaceBlanks(sheetName As String, rge As String)
    Sheets(sheetName).Select
    Range(rge).Select
    Selection.Replace What:="", Replacement:="0", LookAt:=xlWhole, _
        SearchOrder:=xlByRows, MatchCase:=True, SearchFormat:=False, _
        ReplaceFormat:=False
    Range("A1").Select
End Sub
