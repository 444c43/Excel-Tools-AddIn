VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Validations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public ValidationErrors As Collection

Public Sub SetupObject()
    Set ValidationErrors = New Collection
End Sub

Public Sub ValidateSheetNames(expected_names)
    Dim item
    For Each item In expected_names
        If WorkbookHasSheet(CStr(item)) = False Then
            ValidationErrors.Add "Sheet: " & CStr(item) & " missing or incorrect."
        End If
    Next item
End Sub
Private Function WorkbookHasSheet(sheet_name$) As Boolean
    Dim sht As Worksheet
    WorkbookHasSheet = False
    For Each sht In ActiveWorkbook.Sheets
        If sht.name = sheet_name Then
            WorkbookHasSheet = True
            Exit Function
        End If
    Next sht
End Function

Sub ValidateHeaders(expected_headers)
    Dim item
    For Each item In expected_headers
        If SheetsContainHeaders(CStr(item)) = False Then
            ValidationErrors.Add "Header: " & CStr(item) & " missing or incorrect."
        End If
    Next item
End Sub
Private Function SheetsContainHeaders(header_name$) As Boolean
    Dim sht As Worksheet
    SheetsContainHeaders = False
    For Each sht In ActiveWorkbook.Sheets
        If sht.Range("A1").Value = header_name Then
            SheetsContainHeaders = True
            Exit Function
        End If
    Next sht
End Function

Sub ValidateCustomerData()
    
End Sub

