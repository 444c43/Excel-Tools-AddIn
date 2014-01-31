Attribute VB_Name = "AS400Labels"
Option Explicit
Private AS400Labels As SheetSetups
Private ShtNms()
Private ShtNm()
Private HdrNm()
Private ShtHeaders()
Private DateVal As String

Sub HARRYO2Z()
    Set AS400Labels = New SheetSetups
    DateVal = Calculations.DDMMMYY(Day(Now))
    ShtNms = Array("HARRYO2Z " & DateVal)
    ShtHeaders = Array("HARRYO2Z", "Data 1", "Data 2")
    Call SetupActions(ShtNms)
    Call LBLHeaders(ShtHeaders)
End Sub
Sub THIRDPARTY()
    Set AS400Labels = New SheetSetups
    DateVal = Calculations.DDMMMYY(Day(Now))
    ShtNms = Array("3RDPARTY " & DateVal)
    ShtHeaders = Array("3RDPARTY", "Part #", "Unique ID", "Part Description", "Bin Size", "Bin Color", "Card Qty", _
                        "Location", "Supplier", "Unique ID", "Label Count")
    Call SetupActions(ShtNms)
    Call LBLHeaders(ShtHeaders)
End Sub
Sub SFBUILD()
    Set AS400Labels = New SheetSetups
    DateVal = Calculations.DDMMMYY(Day(Now))
    ShtNms = Array("SFBUILD " & DateVal)
    ShtHeaders = Array("SFBUILD", "Serial #", "Customer #", "Ship To", "Plant Code", "Customer Part Number", "GFC Part Number", _
        "Card Qty", "Pkg Type", "Station Number", "Station Description", "Box Qty", "Line Description", "Bin Sys", "Bin Size", _
        "NO Charge SN", "Customer PO", "Country of Origin", "Revision Level", "Comment # 3", "Future Area", "Skipped?")
    Call SetupActions(ShtNms)
    Call LBLHeaders(ShtHeaders)
End Sub
Sub SFUPDATE()
    Set AS400Labels = New SheetSetups
    DateVal = Calculations.DDMMMYY(Day(Now))
    ShtNms = Array("SFUPDATE " & DateVal)
    ShtHeaders = Array("SFUPDATE", "Serial #", "Customer #", "Ship To", "Plant Code", "Customer Part Number", "GFC Part Number", _
        "Card Qty", "Pkg Type", "Station Number", "Station Description", "Box Qty", "Line Description", "Bin Sys", "Bin Size", _
        "NO Charge SN", "Customer PO", "Country of Origin", "Revision Level", "Comment # 3", "Future Area", "Skipped?")
    Call SetupActions(ShtNms)
    Call LBLHeaders(ShtHeaders)
End Sub
Sub STAND1X3()
    Set AS400Labels = New SheetSetups
    DateVal = Calculations.DDMMMYY(Day(Now))
    ShtNms = Array("STAND1X3 " & DateVal)
    ShtHeaders = Array("STAND1X3", "Serial #", "Print Qty")
    Call SetupActions(ShtNms)
    Call LBLHeaders(ShtHeaders)
End Sub
Sub ONELINE()
    Set AS400Labels = New SheetSetups
    DateVal = Calculations.DDMMMYY(Day(Now))
    ShtNms = Array("ONELINE " & DateVal)
    ShtHeaders = Array("ONELINE", "Data 1", "Data 2")
    Call SetupActions(ShtNms)
    Call LBLHeaders(ShtHeaders)
End Sub

Private Sub SetupActions(SheetNames())
    Application.ScreenUpdating = False
    Call AS400Labels.AdjustSheets(SheetNames)
    Sheets(1).Select
    Application.ScreenUpdating = True
End Sub

Private Sub LBLHeaders(HdrValues())
    Dim x As Byte
    For x = 0 To UBound(HdrValues)
        Range(Cells(1, x + 1), Cells(1, x + 1)).Value = HdrValues(x)
    Next x
    Range("A2").Formula = "=SUM(65535-COUNTBLANK(B:B))+1"
    Range("A1:A2").Font.Bold = True
    Range("A1:A2").Font.ColorIndex = 3
    Range(Cells(1, 2), Cells(1, UBound(HdrValues) + 1)).Font.Bold = True
    Call SheetFormatting.AllCellsFit
End Sub
