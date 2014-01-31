Attribute VB_Name = "Calculations"
Option Explicit

Public Function DDMMMYY(Dy As String)
' returns the date in the following format 01Jan10
Dim Mth() As Variant
Mth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    If Len(Dy) = 1 Then
        Dy = "0" & Dy
    End If
    DDMMMYY = Dy & Mth(Month(Now) - 1) & right(Year(Now), 2)
End Function
