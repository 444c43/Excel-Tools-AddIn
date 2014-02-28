VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "UniqueContent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WSFunc As Object
Private UniqueItemList As RangeByHeader
Private MainTab As RangeByHeader
Private i%, passed_value$

Public UniqueItems As Collection
Public NotOrdered As Collection

'setup procedures
Public Sub Initialize(serial_sheet$, serial_header$)
    Set UniqueItemList = New RangeByHeader
    Set WSFunc = Application.WorksheetFunction
    Set UniqueItems = New Collection
    Set NotOrdered = New Collection
    Call SetSearchColumn(serial_sheet, serial_header)
End Sub
Public Sub SetSearchColumn(serial_sheet$, serial_header$)
    Call UniqueItemList.SetProperties(serial_sheet, serial_header)
End Sub

'procedures to compile lists
Public Sub GetUniqueList()
    For i = 2 To UniqueItemList.LastRow
        passed_value = UniqueItemList.ColumnRange.Range("A" & i).Value
        If CollectionHasItem() = False Then
            Call AddIfCellNotEmpty(i)
        End If
    Next i
End Sub
Public Sub GetNotOrdered()
    For i = 2 To UniqueItemList.LastRow
        If IsOrdered(i) Then
            Call AddIfCellNotEmpty(i)
        End If
    Next i
End Sub
Public Sub GetInactive(sheet_name$)
    Set MainTab = New RangeByHeader
    Call MainTab.SetProperties(sheet_name, "Customer Part #")
    
    For i = 2 To UniqueItemList.LastRow
        passed_value = UniqueItemList.ColumnRange.Range("A" & i).Value
        If CollectionHasItem() = False And PartNotOrdered() = False Then
            Call AddIfCellNotEmpty(i)
        End If
    Next i
End Sub

'supporting functions and procedures
Private Function PartNotOrdered() As Boolean
    Dim x%
    
    PartNotOrdered = False
    For x = 2 To MainTab.LastRow
        If MainTab.ColumnRange.Range("A" & x).Value = passed_value Then
            PartNotOrdered = WasOrdered(x)
        End If
    Next x
End Function
Private Sub AddIfCellNotEmpty(i%)
    If UniqueItemList.ColumnRange.Range("A" & i).Value <> "" Then
        passed_value = UniqueItemList.ColumnRange.Range("A" & i).Value
        If CollectionHasItem() = False Then
            UniqueItems.Add UniqueItemList.ColumnRange.Range("A" & i).Value
        End If
    End If
End Sub
Private Function CollectionHasItem() As Boolean
    Dim current_item As Variant
    CollectionHasItem = False
    For Each current_item In UniqueItems
        If passed_value = current_item Then
            CollectionHasItem = True
            Exit For
        End If
    Next current_item
End Function
Private Function IsOrdered(i%) As Boolean
    IsOrdered = (UniqueItemList.ColumnRange.Range("A" & i).Offset(0, -3).Value = "Not Scanned")
End Function
Private Function WasOrdered(i%) As Boolean
    WasOrdered = (MainTab.ColumnRange.Range("A" & i).Offset(0, -3).Value = "Scanned")
End Function
