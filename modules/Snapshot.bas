Attribute VB_Name = "Snapshot"
Option Explicit
Private x, TotalRows%
Private CustomerInfo As New Customer
Private RowsAndColumns As RowsColsCount
Private SnapshotPage As PageOptions
Private SerialFileWS, SnapshotWS, WSFunct As Object
Private CurrentValues As SnapshotValues
Private SnapshotTextFormat As TextFormat

''============================================================
''  REVIEW Snapshot
''      below code formats snapshot page
''============================================================
Public Sub Run()
    Application.ScreenUpdating = False
        Call InstantiateCustomerObject
        Call RowsAndColumns.GetRowsAndCols(CustomerInfo.AcctNumber)
        Sheets.Add
        Call UpdateWorksheet
        Call InstantiateSheetObjects
        Call AddValues
        Call CreateLayout
        Call AddGraphs
        SnapshotWS.Range("A1").Select
    Application.ScreenUpdating = True
End Sub

Private Sub InstantiateCustomerObject()
    Set CustomerInfo = New Customer
    Call CustomerInfo.SetupSnapshotCustomer(activesheet.name)
    Set RowsAndColumns = New RowsColsCount
    Set CurrentValues = New SnapshotValues
    Set SnapshotTextFormat = New TextFormat
End Sub

Private Sub UpdateWorksheet()
    Set SnapshotPage = New PageOptions
    activesheet.name = "Snapshot"
    Call SnapshotPage.ChangeFooters("", _
        "Date Generated: " & Month(Now()) & "/" & Day(Now()) & "/" & Year(Now()), _
        CustomerInfo.name & " " & Sheets(CustomerInfo.AcctNumber).PageSetup.RightHeader)
    Call SnapshotPage.ChangeOrientation(True, False, xlLandscape)
    Call SnapshotPage.ChangeLeftRightMargins(0.5, 0.5)
    Call SnapshotPage.ChangeTopBottomMargins(0.5, 0.5)
    Call SnapshotPage.ChangeHeaderFooterMargins(0.5, 0.3)
End Sub

Private Sub InstantiateSheetObjects()
    Set SerialFileWS = Application.Sheets(CustomerInfo.AcctNumber)
    Set SnapshotWS = Application.Sheets("Snapshot")
    Set WSFunct = Application.WorksheetFunction
    Set SnapshotPage = New PageOptions
End Sub

Private Sub CreateLayout()
    Call Borders
    Call Headers
    Call Comments
End Sub

Private Sub AddValues()
    Call CurrentValues.SetProperties(CustomerInfo.AcctNumber)
    Call SerialValues
    Call PartValues
    Call LoopValues
    Call LegendValues
End Sub

Private Sub AddGraphs()
    Call BarGraph
    Call PieChartAdd
End Sub

Private Sub Borders()
    Dim CellsToMerge(), ColumnRanges(), RowRanges() As Variant
    Dim ColumnWidths(), RowHeight() As Variant
    Dim SnapshotFormat As WorksheetFormatting
    
    Set SnapshotFormat = New WorksheetFormatting
        
    'Set column and row sizes
    ColumnRanges = Array("A:A,F:F,K:K", "B:B,G:G", "C:C,H:H", "D:E,I:J")
    ColumnWidths = Array(4, 20, 12, 10)
    RowRanges = Array("1:1", "2:2", "4:4,23:23,25:25", "3:3,5:22,25:41")
    RowHeight = Array(24, 18.75, 16.5, 12.75)
    
    Call SnapshotFormat.SetColumnAndRowSizes(ColumnRanges, ColumnWidths, RowRanges, RowHeight)
        
    Call SnapshotFormat.MergeCells("A1:F1,G1:K1,A2:F2,G2:K2")
    Call SnapshotFormat.MergeCells("B4:E4,B25:E25,G4:J4,G23:J23")
    Call SnapshotFormat.MergeCells("H24:J24,H25:J25,H26:J26,H27:J27")
    Call SnapshotFormat.MergeCells("H28:J28,H29:J29,H30:J30,H31:J31")
    
    Range("A1:K2,B4:E4,G4:J4,B25:E25,G23:J23").Select
    Call SnapshotFormat.CellColor
    
    SnapshotFormat.SnapshotBorders
End Sub

Private Sub Headers()
    'major headers
    Call TextInCell(Array("A1", "G1", "A2", "G2", "B4", "B25", "G4", "G23"), _
        Array(18, 18, 14, 14, 12, 12, 12, 12), _
        Array("General Fasteners Customer Review", "Stock WHSE: " & CustomerInfo.ShippingWHSE, _
        CustomerInfo.name & " : " & Sheets(CustomerInfo.AcctNumber).PageSetup.RightHeader, _
        "Delivery Freq: " & CustomerInfo.DeliveryFrequency, "Serial Numbers", _
        "Part Numbers", "Sales and Serial Values", "Legend"))
    
    'sub headers - shared
    Call SetMultipleCellValues( _
        Array("C5,H5,C25", "D5,I5,D25", "E5,J5,E25", "B9,B28"), Array("Current", "Prev 1", "Prev 2", "Total"))
        
    'sub headers - serial numbers
    Call SetMultipleCellValues( _
        Array("B6", "B7", "B8", "B10", "B11"), Array("Scanned", "Not Scanned", "Inactive", "Missing Piece Price", "Wkly Bin Scan Avg"))

    'sub headers - part numbers
    Call SetMultipleCellValues( _
        Array("B26", "B27"), Array("Ordered", "Not Ordered"))
    
    'sub headers - loop value data
    Call SetMultipleCellValues( _
        Array("G6", "G7", "G8"), Array("Sales Value", "Not Scanned Value", "Inactive Value"))

    'headers, bold (top)
    SnapshotWS.Range("C5:E5,H5:J5,C25:E25").Font.Bold = True
    SnapshotWS.Range("C5:E5,H5:J5,C25:E25").Font.Italic = True
    SnapshotWS.Range("C5:E5,H5:J5,C25:E25").HorizontalAlignment = xlCenter
    'headers, italic (left)
    SnapshotWS.Range("B6:B11,G6:G8,B26:B28,G24:G31").Font.Italic = True
    SnapshotWS.Range("B6:B11,G6:G8,B26:B28,G24:G31").HorizontalAlignment = xlRight
    'items in blue to denote sum total
    SnapshotWS.Range("B6:B8,B26,B27").Font.ColorIndex = 5
    'sales and serial values header colors
    SnapshotWS.Range("H5").Interior.ColorIndex = 17
    SnapshotWS.Range("I5").Interior.ColorIndex = 18
    SnapshotWS.Range("J5").Interior.ColorIndex = 19
End Sub

Private Sub TextInCell(rge As Variant, FontS As Variant, FontV As Variant)
    Dim i As Byte
    For i = 0 To UBound(rge)
        Range(rge(i)).Select
        With Selection
            .Font.Size = FontS(i)
            .Font.ColorIndex = 2
            .Font.Bold = True
            .Value = FontV(i)
        End With
    Next i
End Sub

Private Sub SetMultipleCellValues(cell_rge As Variant, cell_values As Variant)
    Dim i As Byte
    
    For i = 0 To UBound(cell_rge)
        Range(cell_rge(i)).Value = cell_values(i)
    Next i
End Sub

Private Sub Comments()
    Dim ShapeObj As Object
    Set ShapeObj = activesheet.Shapes
    Range("A1").Select
    With ShapeObj.AddShape(Type:=msoShapeRectangle, left:=25.5, top:=432, Width:=600, Height:=110)
        .name = "Comments"
    End With
    
    ShapeObj("Comments").Select
    ' format text for Legend textbox
    Selection.Characters.Text = "Comments: "
    With Selection.Characters(Start:=1, Length:=9).Font
        .name = "Arial"
        .FontStyle = "Bold"
        .Size = 10
    End With
End Sub

Private Sub SerialValues()
    'serial numbers
    SnapshotWS.Range("C6").Value = CurrentValues.Scanned
    SnapshotWS.Range("C7").Value = CurrentValues.NotScanned
    SnapshotWS.Range("C8").Value = CurrentValues.Inactive
    SnapshotWS.Range("C10").Value = CurrentValues.Missing
    SnapshotWS.Range("C11").Value = CurrentValues.WeeklyAvg
    
    SnapshotWS.Range("C9").Formula = "=SUM(C6:C8)" 'sum formula for serials
    SnapshotWS.Range("C9:E9").FillRight
'    SnapshotWS.Range("E9").Formula = "=SUM(E6:E8)" 'sum formula for serials
End Sub

Private Sub PartValues()
    SnapshotWS.Range("C26").Value = CurrentValues.OrderedParts
    SnapshotWS.Range("C27").Value = CurrentValues.NotOrderedParts

    SnapshotWS.Range("C28").Formula = "=SUM(C26:C27)" 'sum for ordered
    SnapshotWS.Range("C28:E28").FillRight
'    SnapshotWS.Range("D28").Formula = "=SUM(D26:D27)" 'sum for ordered
'    SnapshotWS.Range("E28").Formula = "=SUM(E26:E27)" 'sum for ordered
End Sub
Private Sub LoopValues()
    
    SnapshotWS.Range("H6").Value = CurrentValues.SalesValue
    SnapshotWS.Range("H7").Value = CurrentValues.NotScannedValue
    SnapshotWS.Range("H8").Formula = CurrentValues.InactiveValue
    SnapshotWS.Range("H6:J8").Select
    Call SnapshotTextFormat.NmrFmt("$#,##0.00", xlRight)
End Sub

Private Sub LegendValues()
Dim LegendLeft() As Variant
Dim LegendRight() As Variant
    
LegendLeft = Array("Scanned", "Not Scanned", "Inactive", "Total", _
            "Missing Piece Price", "Scanned Value", "Not Scanned Value", "Inactive Value")
LegendRight = Array("replenished serial numbers", "non-replenished serial numbers", "not scanned or replenished 1 year+", _
            "sum of preceding items in blue", "serials missing piece price", "sales only value", _
            "serial file value of serials not scanned", "serial file value of inactive serials")
For x = 24 To 31
    SnapshotWS.Range("G" & x).Value = LegendLeft(x - 24)
    SnapshotWS.Range("H" & x).Value = LegendRight(x - 24)
Next x
End Sub
Private Sub BarGraph()
    With Charts.Add
        .ChartType = xlColumnClustered
        .SetSourceData Source:=Sheets("Snapshot").Range("G6:J8"), PlotBy:=xlColumns
        .location Where:=xlLocationAsObject, name:="Snapshot"
    End With
    
    ActiveChart.HasLegend = False
    ActiveChart.HasDataTable = False

    With ActiveChart.Parent
         .Height = 162 ' size
         .Width = 288  ' size
         .top = 131  ' position
         .left = 338   ' position
         .name = "Loop Value"
     End With

    activesheet.ChartObjects("Loop Value").Select
    ActiveChart.ChartArea.Select
    ActiveChart.Axes(xlCategory).Select
    Selection.TickLabels.AutoScaleFont = False
    Call SnapshotTextFormat.ObjectFont(Selection.TickLabels.Font, "Arial", "Regular", 10, xlAutomatic)
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabels.AutoScaleFont = False
    Call SnapshotTextFormat.ObjectFont(Selection.TickLabels.Font, "Arial", "Regular", 10, xlAutomatic)
End Sub
Private Sub PieChartAdd()
    Dim myChtObj As ChartObject
    
    Set myChtObj = activesheet.ChartObjects.Add(left:=25, Width:=288, top:=175, Height:=140)
    
    With myChtObj.Chart
        .ChartType = xlPie
        .SetSourceData Source:=Sheets("Snapshot").Range("B6:C8"), PlotBy:=xlColumns
        .location Where:=xlLocationAsObject, name:="Snapshot"
    End With
    
    With myChtObj.Chart
        .Parent.name = "Serial Data"
        .HasTitle = True
        'title
        .ChartTitle.Characters.Text = "Current" & Chr(10) & "Serial" & Chr(10) & "Numbers"
        .ChartTitle.AutoScaleFont = False
        .ChartTitle.Font.name = "Arial"
        .ChartTitle.Font.FontStyle = "Bold"
        .ChartTitle.Font.Size = 10
        .ChartTitle.top = 50
        .ChartTitle.left = 0
        'legend
        .Legend.AutoScaleFont = False
        .Legend.Font.name = "Arial"
        .Legend.Font.FontStyle = "Regular"
        .Legend.Font.Size = 10
        .Legend.top = 38
        .Legend.left = 186
        .Legend.Width = 92
        .Legend.Height = 64
        'pie chart properties
        .SeriesCollection(1).Explosion = 14
        .PlotArea.top = 20
        .PlotArea.left = 74
        .PlotArea.Width = 90
        .PlotArea.Height = 90
        .PlotArea.Border.Weight = xlThin
        .PlotArea.Border.LineStyle = xlNone
        .PlotArea.Interior.ColorIndex = xlNone
        .SeriesCollection(1).ApplyDataLabels AutoText:=True, HasLeaderLines:=True, ShowPercentage:=True
        .SeriesCollection(1).DataLabels.AutoScaleFont = False
        .SeriesCollection(1).DataLabels.Font.name = "Arial"
        .SeriesCollection(1).DataLabels.Font.FontStyle = "Regular"
        .SeriesCollection(1).DataLabels.Font.Size = 10
    End With
End Sub

