Attribute VB_Name = "TEST"
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
    For Each name In Array("Serial File", "Review Data", "Review Data")
        On Error GoTo thing
        sheets("Serial File").Select
        If CheckSheetHeaders(CStr(name)) = False Then
            FailedSheetHeaders.Add name
        End If
thing:
        FailedSheetNames.Add CStr(name)
    Next name
End Sub

Private Function CheckSheetHeaders(name$) As Boolean
    Dim header As Variant
    CheckSheetHeaders = False
    For Each header In Array("GFCSR#", "SERIAL", "CONO80")
        If name = header Then
            CheckSheetHeaders = True
            Exit Function
        End If
    Next header
End Function


Sub Test()
MenuSetup.DestroyToolbarMenu
End Sub










