Attribute VB_Name = "Calculations"
Option Explicit

Public Function DDMMMYY(current_day As String)
    ' returns the date in the following format 01Jan10
    Dim Mth() As Variant
    Mth = Array("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")
    If Len(current_day) = 1 Then
        current_day = "0" & current_day
    End If
    DDMMMYY = current_day & Mth(Month(Now) - 1) & right(Year(Now), 2)
End Function
