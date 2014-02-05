Attribute VB_Name = "Snapshot"
Option Explicit
Private TotalRows As Integer
Private x As Integer
Private CustomerInfo As New Customer
Private SerialFileWS As Object
Private SnapshotWS As Object
Private WSFunct As Object

''============================================================
''  REVIEW Snapshot
''      below code formats snapshot page
''============================================================
Public Sub Run()
    Call CustomerObject
    Application.ScreenUpdating = False
    Call SheetAdd
    'select sheet
    Sheets("Snapshot").Select
    'set sheets and worksheet function as object (reduce typing)
    Set SerialFileWS = Application.Sheets(CustomerInfo.AcctNumber)
    Set SnapshotWS = Application.Sheets("Snapshot")
    Set WSFunct = Application.WorksheetFunction
    TotalRows = SerialFileWS.Range("A36000").End(xlUp).Row
    'call subs
    Borders
    Headers
    SerialValues
    PartValues
    LoopValues
    LegendValues
    BarGraph
    PieChartAdd
    Comments
    'remove price data
    'SerialFileWS.Range("R:U").Delete
    SnapshotWS.Range("A1").Select
    'update screen
    Application.ScreenUpdating = True
End Sub
Private Sub CustomerObject()
    Set CustomerInfo = New Customer
    Call CustomerInfo.SetupCustomerData(activesheet.name)
End Sub
Public Function SheetAdd()
    Sheets.Add
    With activesheet 'mySht
        .name = "Snapshot"
        .PageSetup.LeftFooter = "Date Generated: " & Month(Now()) & "/" & Day(Now()) & "/" & Year(Now())
        .PageSetup.RightFooter = CustomerInfo.name & " " & Sheets(CustomerInfo.AcctNumber).PageSetup.RightHeader
        .PageSetup.CenterHorizontally = True
        .PageSetup.orientation = xlLandscape
        .PageSetup.LeftMargin = Application.InchesToPoints(0.5)
        .PageSetup.RightMargin = Application.InchesToPoints(0.5)
        .PageSetup.TopMargin = Application.InchesToPoints(0.5)
        .PageSetup.BottomMargin = Application.InchesToPoints(0.5)
        .PageSetup.HeaderMargin = Application.InchesToPoints(0.5)
        .PageSetup.FooterMargin = Application.InchesToPoints(0.3)
    End With
End Function
Sub Borders()
    'set column widths to match text
    SnapshotWS.Range("A:A,F:F,K:K").ColumnWidth = 4
    SnapshotWS.Range("B:B,G:G").ColumnWidth = 20
    SnapshotWS.Range("C:C,H:H").ColumnWidth = 12
    SnapshotWS.Range("D:E,I:J").ColumnWidth = 10
    
    'set row height to match text
    Rows("1:1").RowHeight = 24
    Rows("2:2").RowHeight = 18.75
    Range("4:4,23:23,24:24").RowHeight = 16.5
    Range("3:3,5:22,25:41").RowHeight = 12.75
    
    'merge primary and secondary header cells
    'VMI Customer, Stock WHSE, Period, Del Freq
    Call SheetFormatting.MergeCells("A1:F1,G1:K1,A2:F2,G2:K2")
    
    'serial & part numbers, sales, legend
    Call SheetFormatting.MergeCells("B4:E4,B24:E24,G4:J4,G23:J23")
    
    'right side legend merges
    Call SheetFormatting.MergeCells("H24:J24,H25:J25,H26:J26,H27:J27")
    Call SheetFormatting.MergeCells("H28:J28,H29:J29,H30:J30,H31:J31")
    
    'inside color
    SnapshotWS.Range("A1:K2,B4:E4,G4:J4,B24:E24,G23:J23").Select
    Call SheetFormatting.CellColor
    
    'border edges
    Call SheetFormatting.BorderEdges("A1:K2,B4:E10,G4:J8,B24:E28,G23:J31", xlContinuous, xlThick, xlAutomatic)
    Call SheetFormatting.BorderEdges("A1:K41", xlContinuous, xlThick, xlAutomatic)
   
    'border insides
    SnapshotWS.Range("B4:E10,G4:J8,B24:E28,G23:J31").Select
    Call SheetFormatting.BorderInside("B4:E10,G4:J8,B24:E28,G23:J31", xlContinuous, xlThin, xlAutomatic)
End Sub
Sub Headers()
    'major headers
    Call TextFormatting.TextInCell("A1", 18, 2, True, "General Fasteners Customer Review")
    Call TextFormatting.TextInCell("A2", 14, 2, True, CustomerInfo.name & " : " & Sheets(CustomerInfo.AcctNumber).PageSetup.RightHeader)
    Call TextFormatting.TextInCell("G1", 18, 2, True, "Stock WHSE: " & CustomerInfo.ShippingWHSE)
    Call TextFormatting.TextInCell("G2", 14, 2, True, "Delivery Freq: " & CustomerInfo.DeliveryFrequency)
    Call TextFormatting.TextInCell("B4", 12, 2, True, "Serial Numbers")
    Call TextFormatting.TextInCell("B24", 12, 2, True, "Part Numbers")
    Call TextFormatting.TextInCell("G4", 12, 2, True, "Sales and Serial Values")
    Call TextFormatting.TextInCell("G23", 12, 2, True, "Legend")
    'sub headers - shared
    SnapshotWS.Range("C5,H5,C25").value = "Current"
    SnapshotWS.Range("D5,I5,D25").value = "Prev 1"
    SnapshotWS.Range("E5,J5,E25").value = "Prev 2"
    SnapshotWS.Range("B9,B28").value = "Total"
    'sub headers - serial numbers
    SnapshotWS.Range("B6").value = "Scanned"
    SnapshotWS.Range("B7").value = "Not Scanned"
    SnapshotWS.Range("B8").value = "Inactive"
    SnapshotWS.Range("B10").value = "Missing Piece Price"
    'sub headers - part numbers
    SnapshotWS.Range("B26").value = "Ordered"
    SnapshotWS.Range("B27").value = "Not Ordered"
    'sub headers - loop value data
    SnapshotWS.Range("G6").value = "Sales Value"
    SnapshotWS.Range("G7").value = "Not Scanned Value"
    SnapshotWS.Range("G8").value = "Inactive Value"
    'headers, bold (top)
    SnapshotWS.Range("C5:E5,H5:J5,C25:E25").Font.Bold = True
    SnapshotWS.Range("C5:E5,H5:J5,C25:E25").Font.Italic = True
    SnapshotWS.Range("C5:E5,H5:J5,C25:E25").HorizontalAlignment = xlCenter
    'headers, italic (left)
    SnapshotWS.Range("B6:B10,G6:G8,B26:B28,G24:G31").Font.Italic = True
    SnapshotWS.Range("B6:B10,G6:G8,B26:B28,G24:G31").HorizontalAlignment = xlRight
    'items in blue to denote sum total
    SnapshotWS.Range("B6:B8,B26,B27").Font.ColorIndex = 5
    'sales and serial values header colors
    SnapshotWS.Range("H5").Interior.ColorIndex = 17
    SnapshotWS.Range("I5").Interior.ColorIndex = 18
    SnapshotWS.Range("J5").Interior.ColorIndex = 19
End Sub
Sub SerialValues()
    Dim Gx, Hx As Integer
    Dim Px As String
    
    'serial numbers
    SnapshotWS.Range("C6").value = WSFunct.CountIf(SerialFileWS.Range("V:V"), "Scanned") 'scanned
    SnapshotWS.Range("C7").value = WSFunct.CountIf(SerialFileWS.Range("V:V"), "Not Scanned") 'not scanned
    SnapshotWS.Range("C8").value = WSFunct.CountIf(SerialFileWS.Range("V:V"), "Inactive") 'inactive
    
    SnapshotWS.Range("C9").Formula = "=SUM(C6:C8)" 'sum formula for serials
    SnapshotWS.Range("D9").Formula = "=SUM(D6:D8)" 'sum formula for serials
    SnapshotWS.Range("E9").Formula = "=SUM(E6:E8)" 'sum formula for serials
    
    SnapshotWS.Range("C10").value = WSFunct.CountIf(SerialFileWS.Range("R:R"), 0) 'missing pc price
End Sub
Sub PartValues()
Dim ScannedParts, UnscannedParts As Integer
ScannedParts = 0
UnscannedParts = 0
For x = 2 To TotalRows
    If WSFunct.CountIf(SerialFileWS.Range("C" & x & ":" & "C" & TotalRows), SerialFileWS.Range("C" & x)) = 1 Then
        ScannedParts = ScannedParts + 1
        If WSFunct.SumIf(SerialFileWS.Range("C:C"), SerialFileWS.Range("C" & x), SerialFileWS.Range("G:G")) = 0 Then
            UnscannedParts = UnscannedParts + 1
        End If
    End If
Next x
SnapshotWS.Range("C26").value = ScannedParts - UnscannedParts 'ordered
SnapshotWS.Range("C27").value = UnscannedParts 'not ordered
SnapshotWS.Range("C28").Formula = "=SUM(C26:C27)" 'sum for ordered
SnapshotWS.Range("D28").Formula = "=SUM(D26:D27)" 'sum for ordered
SnapshotWS.Range("E28").Formula = "=SUM(E26:E27)" 'sum for ordered
End Sub
Sub LoopValues()
Dim LVScan As Currency
Dim LVUnscan As Currency
Dim LVIn As Currency

LVScan = 0 'sales value
LVUnscan = 0 'not scanned value
LVIn = 0 'inactive value

For x = 2 To TotalRows
    If SerialFileWS.Range("G" & x).value > 0 Then
        LVScan = LVScan + SerialFileWS.Range("G" & x).Value2 * SerialFileWS.Range("R" & x).Value2
    ElseIf SerialFileWS.Range("G" & x).value = 0 And SerialFileWS.Range("P" & x).value <> "I" Then
        LVUnscan = LVUnscan + SerialFileWS.Range("E" & x).Value2 * SerialFileWS.Range("H" & x).Value2 * SerialFileWS.Range("R" & x).Value2
    ElseIf SerialFileWS.Range("G" & x).value = 0 And SerialFileWS.Range("P" & x).value = "I" Then
        LVIn = LVIn + SerialFileWS.Range("E" & x).Value2 * SerialFileWS.Range("H" & x).Value2 * SerialFileWS.Range("R" & x).Value2
    End If
Next x

SnapshotWS.Range("H6").value = LVScan
SnapshotWS.Range("H7").value = LVUnscan
SnapshotWS.Range("H8").Formula = LVIn
SnapshotWS.Range("H6:J8").Select
Call TextFormatting.NmrFmt("$#,##0.00", xlRight)
End Sub

Sub LegendValues()
Dim LegendLeft() As Variant
Dim LegendRight() As Variant
    
LegendLeft = Array("Scanned", "Not Scanned", "Inactive", "Total", _
            "Missing Piece Price", "Scanned Value", "Not Scanned Value", "Inactive Value")
LegendRight = Array("replenished serial numbers", "non-replenished serial numbers", "not scanned or replenished 1 year+", _
            "sum of preceding items in blue", "serials missing piece price", "sales only value", _
            "serial file value of serials not scanned", "serial file value of inactive serials")
For x = 24 To 31
    SnapshotWS.Range("G" & x).value = LegendLeft(x - 24)
    SnapshotWS.Range("H" & x).value = LegendRight(x - 24)
Next x
End Sub
Sub BarGraph()
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
    Call TextFormatting.ObjectFont(Selection.TickLabels.Font, "Arial", "Regular", 10, xlAutomatic)
    ActiveChart.Axes(xlValue).Select
    Selection.TickLabels.AutoScaleFont = False
    Call TextFormatting.ObjectFont(Selection.TickLabels.Font, "Arial", "Regular", 10, xlAutomatic)
End Sub
Sub PieChartAdd()
    Dim myChtObj As ChartObject
    
    Set myChtObj = activesheet.ChartObjects.Add(left:=25, Width:=288, top:=160, Height:=140)
    
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
Sub Comments()
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


