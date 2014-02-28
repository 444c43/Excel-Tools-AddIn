VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ImpExpSnapshot"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Snapshot As Worksheet
Private SnapshotData As Collection
Private FileNames() As Variant

Public Sub Export()
    If CheckForSnapshot Then
        Dim SaveAsName$
        SaveAsName = Sheets("Snapshot").PageSetup.RightFooter
        Call SaveSnapshotData(SaveAsName)
    Else
        MsgBox ("No Snapshot Available!")
    End If
End Sub

Public Sub Import()
    If CheckForSnapshot Then
        Dim location$
        On Error GoTo ErrorControl
            Call GetAndStoreFilenames
            location = FileNames(1)
            Call ImportFile(location)
            Call InsertDataToSnapshot
    Else
        MsgBox ("No Snapshot Available!")
    End If
ErrorControl:
End Sub

Private Sub SaveSnapshotData(filename$)
    Dim SaveToFile
    Set Snapshot = Worksheets("Snapshot")
    'CHANGE THE FILE LOCATION
    'SaveToFile = "C:\Users\dl_2\Desktop\" & filename & ".txt"
    SaveToFile = "T:\Repository\Program Management\Z-Review Data" & filename & ".txt"
    Close #1
    Open SaveToFile For Output As #1
        'current serials
        Call PrintArrayToFile(Array("C6", "C7", "C8", "C10", "C11"))
        'previous serials
        Call PrintArrayToFile(Array("D6", "D7", "D8", "D10", "D11"))
        'current parts
        Call PrintArrayToFile(Array("C26", "C27"))
        'previous parts
        Call PrintArrayToFile(Array("D26", "D27"))
        'current pricing
        Call PrintArrayToFile(Array("H6", "H7", "H8"))
        'previous pricing
        Call PrintArrayToFile(Array("I6", "I7", "I8"))
    Close #1
    MsgBox ("Saved as " & SaveToFile)
End Sub
Private Sub PrintArrayToFile(ranges_arry)
    Dim element As Variant
    For Each element In ranges_arry
        Print #1, Snapshot.Range(element)
    Next element
End Sub

Private Sub GetAndStoreFilenames()
    Dim i%
    Dim dummyString$
 
    FileNames = GetFilenames
 
     ' check for empty array
     On Error Resume Next
     dummyString = FileNames(1)
    
     If Len(dummyString) = 0 Then
         Exit Sub
     End If
     On Error GoTo 0
End Sub

Private Function GetFilenames(Optional title$ = "Select File") As Variant()
  On Error Resume Next
  GetFilenames = Application.GetOpenFilename("TXT (*.txt*), *.txt*", , title, , True)
 End Function

Private Sub ImportFile(location As Variant)
    Dim FileNum%
    Dim FileData$
    
    Set SnapshotData = New Collection
    
    FileNum = FreeFile()
    Open location For Input As #FileNum
        While Not EOF(FileNum)
            Line Input #FileNum, FileData
            SnapshotData.Add (FileData)
        Wend
    Close #FileNum
End Sub

Private Function CheckForSnapshot() As Boolean
    On Error GoTo ErrorHandler
    Sheets("Snapshot").Select
    CheckForSnapshot = True
    Exit Function
    
ErrorHandler:
    CheckForSnapshot = False
End Function

Private Sub InsertDataToSnapshot()
    Dim SaveToFile
    Set Snapshot = Worksheets("Snapshot")
        'current serials
        Call SetValuesFromFile(Array("D6", "D7", "D8", "D10", "D11"), Array(1, 2, 3, 4, 5))
        'previous serials
        Call SetValuesFromFile(Array("E6", "E7", "E8", "E10", "E11"), Array(6, 7, 8, 9, 10))
        'current parts
        Call SetValuesFromFile(Array("D26", "D27"), Array(11, 12))
        'previous parts
        Call SetValuesFromFile(Array("E26", "E27"), Array(13, 14))
        'current pricing
        Call SetValuesFromFile(Array("I6", "I7", "I8"), Array(15, 16, 17))
        'previous pricing
        Call SetValuesFromFile(Array("J6", "J7", "J8"), Array(18, 19, 20))
End Sub
Private Sub SetValuesFromFile(ranges_arry, collection_element)
    Dim count%, i%
    For i = 0 To UBound(ranges_arry)
        Snapshot.Range(ranges_arry(i)).Value = SnapshotData(collection_element(i))
    Next i
End Sub
