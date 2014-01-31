VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "SheetValidations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ErrorList As Collection
Private iterate As Byte

Public Sub ValidateSheets(SheetNames, HeaderValues)
    Set ErrorList = New Collection
    
    If ValidateSheetCount(UBound(SheetNames) + 1) = False Then
        ErrorList.Add ("Incorrect number of sheets!")
    End If

    If ValidateNames(SheetNames) = False Then
        ErrorList.Add ("Mismatched sheet names, or incorrect order!")
    End If
        
    If ValidateHeaders(SheetNames, HeaderValues) = False Then
        ErrorList.Add ("Incorrect data in one or more sheets!")
    End If
End Sub

Private Function ValidateSheetCount(count As Integer) As Boolean
    ValidateSheetCount = (Sheets.count = count)
End Function

Private Function ValidateNames(Names) As Boolean
    ValidateNames = True

    For iterate = 1 To UBound(Names) + 1
        If Sheets(iterate).name <> Names(iterate - 1) Then
            ValidateNames = False
            Exit For
        End If
    Next iterate
End Function

Private Function ValidateHeaders(Names, Headers) As Boolean
    Dim member As Variant
    
    ValidateHeaders = True
    
    For iterate = 0 To UBound(Names)
        If Sheets(Names(iterate)).Range("A1").Value <> Headers(iterate) Then
            ValidateHeaders = False
        End If
    Next
End Function


Private Function WorkbookSheets() As Collection
    Set WorkbookSheets = New Collection

    For iterate = 1 To Sheets.count
        WorkbookSheets.Add Sheets(iterate).name
    Next iterate
End Function

