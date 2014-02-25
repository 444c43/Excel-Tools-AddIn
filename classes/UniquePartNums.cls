VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UniquePartNums"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WSFunc As Object
Private SerialPartList As RangeByHeader
Private ReviewPartList As RangeByHeader

Public UniqueParts As Collection
Public NotOrdered As Collection

Public Sub Initialize(serial_sheet$, serial_header$, review_sheet$, review_header$)
    Call SetObjects
    
    Call SerialPartList.SetProperties(serial_sheet, serial_header)
    Call ReviewPartList.SetProperties(review_sheet, review_header)
End Sub

Public Sub SeparateNotOrdered()
    Dim i%
    For i = 1 To UniqueParts.count
        If WSFunc.CountIf(ReviewPartList.ColumnRange.Range("A:A"), UniqueParts(i)) = 0 Then
            NotOrdered.Add UniqueParts(i)
        End If
    Next i
End Sub

Private Sub SetObjects()
    Set SerialPartList = New RangeByHeader
    Set ReviewPartList = New RangeByHeader
    
    Set WSFunc = Application.WorksheetFunction
    Set UniqueParts = New Collection
    Set NotOrdered = New Collection
End Sub

Public Sub IterateAndAddUnique()
    Dim i%
    
    For i = 2 To SerialPartList.LastRow
        If CountIfValue(i) = 1 Then
            Call AddIfCellNotEmpty(i)
        End If
    Next i
End Sub
Private Sub AddIfCellNotEmpty(i%)
    If SerialPartList.ColumnRange.Range("A" & i).Value <> "" Then
        UniqueParts.Add SerialPartList.ColumnRange.Range("A" & i).Value
    End If
End Sub

Private Function CountIfValue%(i%)
    CountIfValue% = WSFunc.CountIf( _
        SerialPartList.ColumnRange.Range("A" & i & ":A" & SerialPartList.LastRow), _
        SerialPartList.ColumnRange.Range("A" & i))
End Function
