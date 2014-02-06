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

Public Sub ValidateSheets(sheetnames, HeaderValues)
    Set ErrorList = New Collection
    
    If ValidateSheetCount(UBound(sheetnames) + 1) = False Then
        ErrorList.Add ("Incorrect number of sheets!")
    End If

    If ValidateNames(sheetnames) = False Then
        ErrorList.Add ("Mismatched sheet names, or incorrect order!")
    End If
        
    If ValidateHeaders(sheetnames, HeaderValues) = False Then
        ErrorList.Add ("Incorrect data in one or more sheets!")
    End If
End Sub

Private Function ValidateSheetCount(count%) As Boolean
    ValidateSheetCount = (sheets.count = count)
End Function

Private Function ValidateNames(Names) As Boolean
    ValidateNames = True

    For iterate = 1 To UBound(Names) + 1
        If sheets(iterate).name <> Names(iterate - 1) Then
            ValidateNames = False
            Exit For
        End If
    Next iterate
End Function

Private Function ValidateHeaders(Names, Headers) As Boolean
    Dim member As Variant
    
    ValidateHeaders = True
    
    For iterate = 0 To UBound(Names)
        If sheets(Names(iterate)).Range("A1").Value <> Headers(iterate) Then
            ValidateHeaders = False
        End If
    Next
End Function


Private Function WorkbookSheets() As Collection
    Set WorkbookSheets = New Collection

    For iterate = 1 To sheets.count
        WorkbookSheets.Add sheets(iterate).name
    Next iterate
End Function

