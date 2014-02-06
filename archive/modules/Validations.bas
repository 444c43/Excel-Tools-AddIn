Attribute VB_Name = "Validations"
Option Explicit

Public FailedSheetNames As Collection
Public FailedSheetHeaders As Collection

Sub ValidateSheetNames()
    Set FailedSheetNames = New Collection
    Set FailedSheetHeaders = New Collection
    
    Call SheetHeadersToTest
End Sub

Private Sub SheetHeadersToTest()
    Dim name As Variant
    For Each name In Array("Serial File", "Review Data", "Price List")
        On Error GoTo AddFailedSheet
        sheets(name).Select
        If CheckSheetHeaders() = False Then
            FailedSheetHeaders.Add "Bad data on " & name
        End If
    Next name
    Exit Sub
    
AddFailedSheet:
        FailedSheetNames.Add "No sheet named: " & name
        Resume Next
End Sub

Private Function CheckSheetHeaders() As Boolean
    Dim header As Variant
    CheckSheetHeaders = False
    For Each header In Array("GFCSR#", "SERIAL", "CONO80")
        If Range("A1").Value = header Then
            CheckSheetHeaders = True
            Exit Function
        End If
    Next header
End Function


