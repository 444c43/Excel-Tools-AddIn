VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "AS400Headers"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private HeadersAS400(0 To 38) As String
Private HeadersActual(0 To 38) As String

Public Function Convert(sheet_name$)
    Dim TotalColumns As Byte, x As Byte, y As Byte
    
    TotalColumns = Range("IV1").End(xlToLeft).Column
    
    Call HeadersAS400Arr
    Call HeadersActualArr
    
    Sheets(sheet_name).Select
    
    For x = 1 To TotalColumns
        For y = 0 To 38
            Call EvaluateHeader(x, y)
        Next y
    Next x
End Function

Private Sub EvaluateHeader(x As Byte, y As Byte)
    If HeadersAS400(y) = Range(Cells(1, x), Cells(1, x)).Value Then
        Range(Cells(1, x), Cells(1, x)).Value = HeadersActual(y)
    End If
End Sub

Private Sub HeadersAS400Arr()
'AS400 header titles
    HeadersAS400(0) = "CONO80"
    HeadersAS400(1) = "LIST80"
    HeadersAS400(2) = "CATN80"
    HeadersAS400(3) = "ALTI38"
    HeadersAS400(4) = "PDES35"
    HeadersAS400(5) = "PRCE80"
    HeadersAS400(6) = "DTEF8004"
    HeadersAS400(7) = "DTEF"
    HeadersAS400(8) = "SERIAL"
    HeadersAS400(9) = "GFCCS1"
    HeadersAS400(10) = "GFCCS2"
    HeadersAS400(11) = "GFCPLT"
    HeadersAS400(12) = "GFCCP#"
    HeadersAS400(13) = "GFCGF#"
    HeadersAS400(14) = "GFCQTY"
    HeadersAS400(15) = "GFCPKG"
    HeadersAS400(16) = "GFCSTA"
    HeadersAS400(17) = "GFCSTD"
    HeadersAS400(18) = "GFCDE1"
    HeadersAS400(19) = "GFCDE2"
    HeadersAS400(20) = "GFCLN#"
    HeadersAS400(21) = "GFLDES"
    HeadersAS400(22) = "GFLOC1"
    HeadersAS400(23) = "GFLOC2"
    HeadersAS400(24) = "GFNOCH"
    HeadersAS400(25) = "GFSPO#"
    HeadersAS400(26) = "GFCMT1"
    HeadersAS400(27) = "GFCMT2"
    HeadersAS400(28) = "GFCMT3"
    HeadersAS400(29) = "GFFILL"
    HeadersAS400(30) = "GFSSTS"
    HeadersAS400(31) = "LQTY7001"
    HeadersAS400(32) = "WTSU3503"
    HeadersAS400(33) = "GFCSR#"
    HeadersAS400(34) = "GFSSTS"
    HeadersAS400(35) = "GFLUSR"
    HeadersAS400(36) = "GFLUPD"
    HeadersAS400(37) = "GFLTIM"
    HeadersAS400(38) = "GFATYP"
End Sub
Private Sub HeadersActualArr()
'AS400 converted header titles
    HeadersActual(0) = "Company #"
    HeadersActual(1) = "Price List"
    HeadersActual(2) = "GFC Part #"
    HeadersActual(3) = "Customer Part #"
    HeadersActual(4) = "Part Description"
    HeadersActual(5) = "Price"
    HeadersActual(6) = ""
    HeadersActual(7) = "Date Effective"
    HeadersActual(8) = "Serial #"
    HeadersActual(9) = "Customer #"
    HeadersActual(10) = "Ship To"
    HeadersActual(11) = "Plant Code"
    HeadersActual(12) = "Customer Part #"
    HeadersActual(13) = "GFC Part #"
    HeadersActual(14) = "Bin Qty"
    HeadersActual(15) = "Pkg Type"
    HeadersActual(16) = "Station"
    HeadersActual(17) = "Station Description"
    HeadersActual(18) = "Part Description"
    HeadersActual(19) = ""
    HeadersActual(20) = "Mfg Qty"
    HeadersActual(21) = "Line Description"
    HeadersActual(22) = "Bin Sys"
    HeadersActual(23) = "Bin Size"
    HeadersActual(24) = "No Charge"
    HeadersActual(25) = "PO #"
    HeadersActual(26) = "Country of Origin"
    HeadersActual(27) = "Revision Level"
    HeadersActual(28) = "Comment 3"
    HeadersActual(29) = "Future Area"
    HeadersActual(30) = "Serial Status"
    HeadersActual(31) = "Ship Qty"
    HeadersActual(32) = "Weight"
    HeadersActual(33) = "Serial #"
    HeadersActual(34) = "SF Status"
    HeadersActual(35) = "Last User"
    HeadersActual(36) = "Update Date"
    HeadersActual(37) = "Update Time"
    HeadersActual(38) = "Audit Type"
End Sub

