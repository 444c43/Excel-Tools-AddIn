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
    SaveToFile = "C:\Users\dl_2\Desktop\" & filename & ".txt"
    'SaveToFile = "T:\Repository\Program Management\Z-Review Data" & filename & ".txt"
    Close #1
    Open SaveToFile For Output As #1
    'current serials
    Print #1, Snapshot.Range("C6").Value ' scanned
    Print #1, Snapshot.Range("C7").Value 'notscanned
    Print #1, Snapshot.Range("C8").Value 'inactive
    Print #1, Snapshot.Range("C10").Value 'missing
    'previous serials
    Print #1, Snapshot.Range("D6").Value 'scanned
    Print #1, Snapshot.Range("D7").Value 'notscanned
    Print #1, Snapshot.Range("D8").Value 'inactive
    Print #1, Snapshot.Range("D10").Value 'missing
    'current parts
    Print #1, Snapshot.Range("C26").Value 'ordered
    Print #1, Snapshot.Range("C27").Value 'notordered
    'previous parts
    Print #1, Snapshot.Range("D26").Value 'ordered
    Print #1, Snapshot.Range("D27").Value 'notordered
    'current pricing
    Print #1, Snapshot.Range("H6").Value 'salesvalue
    Print #1, Snapshot.Range("H7").Value 'notscannedvalue
    Print #1, Snapshot.Range("H8").Value 'inactivevalue
    'previous pricing
    Print #1, Snapshot.Range("I6").Value 'salesvalue
    Print #1, Snapshot.Range("I7").Value 'notscannedvalue
    Print #1, Snapshot.Range("I8").Value 'inactivevalue
    
    Close #1
    MsgBox ("Saved as " & SaveToFile)
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
    Snapshot.Range("D6").Value = SnapshotData(1)
    Snapshot.Range("D7").Value = SnapshotData(2)
    Snapshot.Range("D8").Value = SnapshotData(3)
    Snapshot.Range("D10").Value = SnapshotData(4)
    'previous serials
    Snapshot.Range("E6").Value = SnapshotData(5)
    Snapshot.Range("E7").Value = SnapshotData(6)
    Snapshot.Range("E8").Value = SnapshotData(7)
    Snapshot.Range("E10").Value = SnapshotData(8)
    'current parts
    Snapshot.Range("D26").Value = SnapshotData(9)
    Snapshot.Range("D27").Value = SnapshotData(10)
    'previous parts
    Snapshot.Range("E26").Value = SnapshotData(11)
    Snapshot.Range("E27").Value = SnapshotData(12)
    'current pricing
    Snapshot.Range("I6").Value = SnapshotData(13)
    Snapshot.Range("I7").Value = SnapshotData(14)
    Snapshot.Range("I8").Value = SnapshotData(15)
    'previous pricing
    Snapshot.Range("J6").Value = SnapshotData(16)
    Snapshot.Range("J7").Value = SnapshotData(17)
    Snapshot.Range("J8").Value = SnapshotData(18)
End Sub
