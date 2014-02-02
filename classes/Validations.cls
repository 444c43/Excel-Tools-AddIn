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
'row and column variables
Private LastColumnPrice%, LastColumnReview
'objects
Private PriceSearchRange, ReviewSearchRange As Range
Private ReviewCustomer, PriceCode As Range
Public ValidationErrors As Collection

Public Sub SetupObject(price_header$, review_header$)
    LastColumnPrice = Sheets("Price List").Range("IV1").End(xlToLeft).Column
    LastColumnReview = Sheets("Review Data").Range("IV1").End(xlToLeft).Column
    
    Set ValidationErrors = New Collection
    Sheets("Price List").Select
    Set PriceSearchRange = Sheets("Price List").Range(Cells(1, 1), Cells(1, LastColumnPrice))
    Sheets("Review Data").Select
    Set ReviewSearchRange = Sheets("Review Data").Range(Cells(1, 1), Cells(1, LastColumnReview))
    Set PriceCode = Sheets("Price List").Range(PriceSearchRange.Find(price_header).Address)
    Set ReviewCustomer = Sheets("Review Data").Range(ReviewSearchRange.Find(review_header).Address)
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
        If sht.Name = sheet_name Then
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
        If sht.Range("A1").value = header_name Then
            SheetsContainHeaders = True
            Exit Function
        End If
    Next sht
End Function

Sub ValidateCustomerData(account_number$, price_code$)
    If account_number <> ReviewCustomer.Cells(2, 1) Then
        ValidationErrors.Add "Serial and Review data do not match!"
    End If
    If price_code <> PriceCode.Cells(2, 1) Then
        ValidationErrors.Add "Incorrect price list for customer!"
    End If
End Sub

