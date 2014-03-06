Attribute VB_Name = "AS400Labels"
Option Explicit
Private SheetNames()
Private SheetHeaders()

Sub HARRYO2Z()
    SheetNames = Array("HARRYO2Z " & Calculations.DDMMMYY(Day(Now)))
    SheetHeaders = Array("HARRYO2Z", "Data 1", "Data 2")
    Call RequestSetup(SheetNames, SheetHeaders)
End Sub
Sub THIRDPARTY()
    SheetNames = Array("3RDPARTY " & Calculations.DDMMMYY(Day(Now)))
    SheetHeaders = Array("3RDPARTY", "Part #", "Unique ID", "Part Description", "Bin Size", "Bin Color", "Card Qty", _
                        "Location", "Supplier", "Unique ID", "Label Count")
    Call RequestSetup(SheetNames, SheetHeaders)
End Sub
Sub SFBUILD()
    SheetNames = Array("SFBUILD " & Calculations.DDMMMYY(Day(Now)))
    SheetHeaders = Array("SFBUILD", "Serial #", "Customer #", "Ship To", "Plant Code", "Customer Part Number", "GFC Part Number", _
        "Card Qty", "Pkg Type", "Station Number", "Station Description", "Box Qty", "Line Description", "Bin Sys", "Bin Size", _
        "NO Charge SN", "Customer PO", "Country of Origin", "Revision Level", "Comment # 3", "Future Area", "Skipped?")
    Call RequestSetup(SheetNames, SheetHeaders)
End Sub
Sub SFUPDATE()
    SheetNames = Array("SFUPDATE " & Calculations.DDMMMYY(Day(Now)))
    SheetHeaders = Array("SFUPDATE", "Serial #", "Customer #", "Ship To", "Plant Code", "Customer Part Number", "GFC Part Number", _
        "Card Qty", "Pkg Type", "Station Number", "Station Description", "Box Qty", "Line Description", "Bin Sys", "Bin Size", _
        "NO Charge SN", "Customer PO", "Country of Origin", "Revision Level", "Comment # 3", "Future Area", "Serial Number Status", _
        "Skipped?")
    Call RequestSetup(SheetNames, SheetHeaders)
End Sub
Sub STAND1X3()
    SheetNames = Array("STAND1X3 " & Calculations.DDMMMYY(Day(Now)))
    SheetHeaders = Array("STAND1X3", "Serial #", "Print Qty")
    Call RequestSetup(SheetNames, SheetHeaders)
End Sub
Sub ONELINE()
    SheetNames = Array("ONELINE " & Calculations.DDMMMYY(Day(Now)))
    SheetHeaders = Array("ONELINE", "Data 1", "Data 2")
    Call RequestSetup(SheetNames, SheetHeaders)
End Sub

Private Sub RequestSetup(SheetNames, SheetHeaders)
    Dim NewRequest As AS400
    Set NewRequest = New AS400
    Call NewRequest.SetupAS400(SheetNames, SheetHeaders)
End Sub

