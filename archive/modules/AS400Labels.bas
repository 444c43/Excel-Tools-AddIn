Attribute VB_Name = "AS400Labels"
Option Explicit
Private AS400Labels As SheetSetup
Private sheetnames()
Private sheetheaders()

' CALLED SUBROUTINES FROM MENUSETUP
Sub HARRYO2Z()
    sheetnames = Array("HARRYO2Z " & Calculations.DDMMMYY(Day(Now)))
    sheetheaders = Array("HARRYO2Z", "Data 1", "Data 2")
    Call SetupAS400(sheetnames, sheetheaders)
End Sub
Sub THIRDPARTY()
    sheetnames = Array("3RDPARTY " & Calculations.DDMMMYY(Day(Now)))
    sheetheaders = Array("3RDPARTY", "Part #", "Unique ID", "Part Description", "Bin Size", "Bin Color", "Card Qty", _
                        "Location", "Supplier", "Unique ID", "Label Count")
    Call SetupAS400(sheetnames, sheetheaders)
End Sub
Sub SFBUILD()
    sheetnames = Array("SFBUILD " & Calculations.DDMMMYY(Day(Now)))
    sheetheaders = Array("SFBUILD", "Serial #", "Customer #", "Ship To", "Plant Code", "Customer Part Number", "GFC Part Number", _
        "Card Qty", "Pkg Type", "Station Number", "Station Description", "Box Qty", "Line Description", "Bin Sys", "Bin Size", _
        "NO Charge SN", "Customer PO", "Country of Origin", "Revision Level", "Comment # 3", "Future Area", "Skipped?")
    Call SetupAS400(sheetnames, sheetheaders)
End Sub
Sub SFUPDATE()
    sheetnames = Array("SFUPDATE " & Calculations.DDMMMYY(Day(Now)))
    sheetheaders = Array("SFUPDATE", "Serial #", "Customer #", "Ship To", "Plant Code", "Customer Part Number", "GFC Part Number", _
        "Card Qty", "Pkg Type", "Station Number", "Station Description", "Box Qty", "Line Description", "Bin Sys", "Bin Size", _
        "NO Charge SN", "Customer PO", "Country of Origin", "Revision Level", "Comment # 3", "Future Area", "Skipped?")
    Call SetupAS400(sheetnames, sheetheaders)
End Sub
Sub STAND1X3()
    sheetnames = Array("STAND1X3 " & Calculations.DDMMMYY(Day(Now)))
    sheetheaders = Array("STAND1X3", "Serial #", "Print Qty")
    Call SetupAS400(sheetnames, sheetheaders)
End Sub
Sub ONELINE()
    sheetnames = Array("ONELINE " & Calculations.DDMMMYY(Day(Now)))
    sheetheaders = Array("ONELINE", "Data 1", "Data 2")
    Call SetupAS400(sheetnames, sheetheaders)
End Sub

'SUBROUTINES THAT SETUP SHEETS
Private Sub SetupAS400(sheets, Headers)
    Set AS400Labels = New SheetSetup
    
    Call SetupActions(sheets)
    Call AddHeaderValues(Headers)
    Call FormatHeaderCells(UBound(Headers) + 1)
End Sub

Private Sub SetupActions(shtnames)
    Application.ScreenUpdating = False
    Call AS400Labels.AdjustSheets(shtnames)
    sheets(1).Select
    Application.ScreenUpdating = True
End Sub

Private Sub AddHeaderValues(HdrValues)
    Dim x As Byte
    For x = 0 To UBound(HdrValues)
        Range(Cells(1, x + 1), Cells(1, x + 1)).Value = HdrValues(x)
    Next x
    
    Range("A2").Formula = "=SUM(65535-COUNTBLANK(B:B))+1"
End Sub

Private Sub FormatHeaderCells(lastCol%)
    Range("A1:A2").Font.Bold = True
    Range("A1:A2").Font.ColorIndex = 3
    Range(Cells(1, 2), Cells(1, lastCol)).Font.Bold = True
    Call SheetFormatting.AllCellsFit
End Sub
