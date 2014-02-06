Attribute VB_Name = "Review"
Option Explicit
Private SubSheet As SubSheets
Private ReviewSheets As SheetSetup
Private SheetValidate As SheetValidations
Private DataValidate As DataValidations
Private ErrorMessage As MessageBox
Public NewCustomer As Customer
Private TotalRows%
Private x%
Public ReviewValidate As ReviewValidations

'ENTRY POINT FOR APPLICATION
Sub Main()
    Application.ScreenUpdating = False
    Call InstantiateObjects
    Call TestAndRun
    Application.ScreenUpdating = True
End Sub
Private Sub InstantiateObjects() 'SETUP OBJECTS
    Set ReviewValidate = New ReviewValidations
    ReviewValidate.ValidateSheetNames   'RUN TESTS
    
    Set NewCustomer = New Customer
    Call NewCustomer.SetupCustomerData(sheets("Serial File").Range("B2").Value)
End Sub
Private Sub TestAndRun()    'CHECK FOR ERROR MESSAGES
    If CheckForErrors(ReviewValidate.FailedSheetNames, ReviewValidate.FailedSheetHeaders) Then
        Call DisplayErrorMessage(ReviewValidate.FailedSheetNames)
        Call DisplayErrorMessage(ReviewValidate.FailedSheetHeaders)
    Else
        If ValidateData(NewCustomer.AcctNumber, NewCustomer.pricecode) Then ' CONFIRM CORRECT CUSTOMER DATA
            Call CallMethods ' CALL METHODS TO RUN ACTUAL REVIEW
        End If
    End If
End Sub
Private Function CheckForErrors(sheetnames As Collection, sheetheaders As Collection) As Boolean
    If sheetnames.count > 0 Or sheetheaders.count > 0 Then
        CheckForErrors = True
        Exit Function
    End If
    CheckForErrors = False
End Function
Private Sub DisplayErrorMessage(item As Collection)
    Dim element As Variant
    For Each element In item
        Debug.Print element
    Next
End Sub
Private Function ValidateData(customernumber$, pricecode$) As Boolean
    Set DataValidate = New DataValidations
    Call DataValidate.ValidateData(customernumber, pricecode)
    ValidateData = EvaluateErrorList(DataValidate.ErrorList)
End Function





' DOCUMENT SETUP
Public Sub SetupReviewSheets()
    Dim sheetnames()
    Set ReviewSheets = New SheetSetup
    sheetnames = Array("Serial File", "Review Data", "Price List")
    Call ReviewSheets.AdjustSheets(sheetnames)
End Sub

Private Function EvaluateErrorList(list As Collection) As Boolean
    Set ErrorMessage = New MessageBox
    
    If list.count > 0 Then
        Call ErrorMessage.MultiLine(vbCritical, "Fix These Errors:", list)
        EvaluateErrorList = False
    Else
        EvaluateErrorList = True
    End If
End Function

Private Sub CreateNewSheet()
    sheets.Add
    activesheet.name = "Inactive Serial Files"
End Sub

Private Sub CallMethods()
    frmReview.Show
    Call Destructors
    Call SetupSubSheetInactiveA
''  set correct sheet focus
    sheets("Serial File").Select
''  set variable TotalRows
    TotalRows = Range("A36000").End(xlUp).Row

    Call TextFormatting.IterateColumns
    Call TextFormatting.HeaderCorrect
    Call ColumnAddRem
    Call HeaderAdjust
    Call FormulaAdd
    Call BordersAdd
    Call CleanUp
    Call SheetFormatting.AllCellsFit
    Call ColumnFormat(TotalRows)
''  set pane freeze
    Call SheetTools.PaneFreeze
    Range("A1").Select
    Call SheetRem
    Call CalculateSerials
    sheets("Serial File").name = NewCustomer.AcctNumber
    sheets(NewCustomer.AcctNumber).PageSetup.LeftHeader = NewCustomer.name
    sheets(NewCustomer.AcctNumber).PageSetup.RightHeader = NewCustomer.ReviewPeriod
    Call SetupSubSheetMissingAndNotOrdered
    sheets(NewCustomer.AcctNumber).Select
    Call Updates
    Call SheetFormatting.AllCellsFit
End Sub

Private Sub Destructors()
    Set ReviewSheets = Nothing
    Set SheetValidate = Nothing
    Set DataValidate = Nothing
    Set ErrorMessage = Nothing
End Sub

Private Sub ColumnAddRem()
    Call SheetFormatting.ColDel("B:C,L:N,S:W", xlToLeft)
    Columns("J:J").Cut
    Call SheetFormatting.ColAdd("F:F", xlToRight)
    Call SheetFormatting.ColAdd("F:G", xlToRight)
    Call SheetFormatting.ColAdd("I:J", xlToRight)
End Sub

Private Sub HeaderAdjust()
    'headers new columns
    Range("F1").Value = "Wkly Avg"
    Range("G1").Value = "Ship Qty"
    Range("I1").Value = "Proposed Bin Sys"
    Range("J1").Value = "Proposed Add/Rem"
    Range("R1").Value = "Pc Price"
    Range("S1").Value = "MLV"
    Range("T1").Value = "Net Chg Value"
    Range("U1").Value = "PLV"
    Range("V1").Value = "Serial Status"
    Range("W1").Value = "To Review"
    Range("X1").Value = "Sales < $15"
    Range("Y1").Value = "Missing Pc Price"
    
    'set all formula columns to general
    Range("F:G,I:J,R:U").NumberFormat = "General"
End Sub
Private Sub FormulaAdd()
    'text to columns serial numbers in Review Data
    sheets("Review Data").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited
    'average bins shipped per week
    Range("F2").FormulaR1C1 = "=ROUNDUP((RC[1]/" & NewCustomer.ReviewWeeks & "),0)"
    Range("G2").FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[-6],'Review Data'!R2C1:R" & sheets("Review Data").Range("A36000").End(xlUp).Row & "C24,24,False))=TRUE,0,VLOOKUP(RC[-6],'Review Data'!R2C1:R" & sheets("Review Data").Range("A36000").End(xlUp).Row & "C24,24,False))"
    Range("I2").FormulaR1C1 = NewCustomer.Formula ' "=ROUNDUP(SUM((RC[-2]/" & WeekCalc & ")/RC[-4]),0)"
    Range("J2").FormulaR1C1 = "=IF(RC[-3]=0,0,SUM(RC[-1]-RC[-2]))"
    Range("R2").FormulaR1C1 = "=VLOOKUP(RC[-14],'Price List'!R2C3:R" & sheets("Price List").Range("A36000").End(xlUp).Row & "C6,4,False)"
    Range("S2").FormulaR1C1 = "=SUM((RC[-1]*RC[-14])*RC[-11])"
    Range("T2").FormulaR1C1 = "=SUM((RC[-2]*RC[-15])*RC[-10])"
    Range("U2").FormulaR1C1 = "=SUM((RC[-3]*RC[-16])*RC[-12])"
''  Copy down
    Range("F2:G" & TotalRows & ",I2:J" & TotalRows & ",R2:U" & TotalRows).FillDown
''    Range("F2").Activate
''    Selection.FillDown
End Sub
Private Sub BordersAdd()
    Call SheetFormatting.BorderEdges("A1:Y" & TotalRows, xlContinuous, xlThin, xlAutomatic)
    Call SheetFormatting.BorderInside("A1:Y" & TotalRows, xlContinuous, xlThin, xlAutomatic)
End Sub
Private Sub CleanUp()
''  Remove #N/A values from Pc Price column
    For x = 2 To TotalRows
        If Range("R" & x).Text = "#N/A" Then
            Range("R" & x).Value = 0
            Range("A" & x).Interior.ColorIndex = 3
            Range("Y" & x).Value = "x"
        End If
    Next x
End Sub
Public Sub ColumnFormat(TRows%)
''  Format entire columns:
    'autofit
    Cells.EntireColumn.AutoFit
    
    'column widths
    Range("C:D,V:V").ColumnWidth = 15
    Range("F:F,R:R,X:X").ColumnWidth = 11
    Range("G:G,Y:Y").ColumnWidth = 8
    Columns("H:H").ColumnWidth = 7
    Range("I:I,W:W").ColumnWidth = 10
    Columns("J:J").ColumnWidth = 12
    
    'first row text format
    Range("A1:Y1").HorizontalAlignment = xlCenter
    Range("A1:Y1").VerticalAlignment = xlCenter
    Range("A1:Y1").WrapText = True
    Range("A1:Y1").Font.Bold = True

    'column C, H, L & O
    Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited
    Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlDelimited
    Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlDelimited
    Columns("O:O").TextToColumns Destination:=Range("O1"), DataType:=xlDelimited
    'center above columns
    Range("A:D,K:M,O:P,V:Y").HorizontalAlignment = xlCenter
    
    'columns F-G, Q, R-T
    Range("F2:G" & TRows).Select
    Call TextFormatting.NmrFmt("#,##0", xlRight)
    Range("R2:R" & TRows).Select
    Call TextFormatting.NmrFmt("$#,##0.00000", xlRight)
    Range("S2:U" & TRows).Select
    Call TextFormatting.NmrFmt("$#,##0.00", xlRight)
    
    'green highlight over review information
    Range("I1:J1").Interior.ColorIndex = 4
    'copy and paste values (formula to values)
    Columns("G:G").Copy
    Range("G1").PasteSpecial Paste:=xlPasteValues
    Columns("R:R").Copy
    Range("R1").PasteSpecial Paste:=xlPasteValues
End Sub

Sub CalculateSerials()
    Dim Gx, Hx%
    Dim Px$

    TotalRows = Range("A36000").End(xlUp).Row
    For x = 2 To TotalRows
        Gx = Range("G" & x).Value
        Hx = Range("H" & x).Value
        Px = Range("P" & x).Value
            
        If Gx > 0 Then
            Range("V" & x).Value = "Scanned"
        End If
        
        'ADD CLARITY HERE IN COMMENTS
        If Gx = 0 And Px <> "I" And Hx > 0 Then
            Range("V" & x).Value = "Not Scanned" 'not scanned
        ElseIf Gx = 0 And Hx = 0 Then
            Range("V" & x).Value = "Inactive Zero Bins" 'zero bins
        ElseIf Gx = 0 And Px = "I" And Hx > 0 Then
            Range("V" & x).Value = "Inactive" 'inactive
        ElseIf Gx > 0 And Px = "I" And Hx > 0 Or Gx > 0 And Px <> "I" And Hx = 0 Or _
            Gx > 0 And Px = "I" And Hx = 0 Or Gx = 0 And Px <> "I" And Hx = 0 Then
                Range("W" & x).Value = "x" 'to review
        End If
        
        'calculate all sales under $15
        If Range("G" & x).Value > 0 And Range("G" & x).Value * Range("R" & x).Value < 15 Then
            Range("X" & x).Value = "x"
        End If
    Next x
End Sub

Private Sub SheetRem()
    Application.DisplayAlerts = False
    sheets("Price List").Delete
    sheets("Review Data").Delete
    Application.DisplayAlerts = True
End Sub

Private Sub SetupSubSheetInactiveA()
    Set SubSheet = New SubSheets
    Call SubSheet.CreateNewSheet("Inactive Serials", "Price List")
    Call SubSheet.MoveDeleted
End Sub

Private Sub SetupSubSheetMissingAndNotOrdered()
    Dim Headers()
    
    Call FormatSubSheet(Array("Customer Number", "GFC Number"), "Missing Pc Price", NewCustomer.AcctNumber)
    Call SubSheet.MissingPcPrice(NewCustomer.AcctNumber)
    
    Call FormatSubSheet(Array("Customer Number", "GFC Number"), "Parts Not Ordered", NewCustomer.AcctNumber)
    Call SubSheet.PartsNotOrdered(NewCustomer.AcctNumber)
End Sub

Private Sub FormatSubSheet(headerdata As Variant, NewSheetName$, SearchSheetName$)
    Call SubSheet.CreateNewSheet(NewSheetName, "Inactive Serials")
    Call SubSheet.SetHeaders(headerdata, "A1:B1")
End Sub

Private Sub Updates()
    Columns("I:I").Insert Shift:=xlToRight
    Range("I1").Value = "Avg Scans"
    Selection.NumberFormat = "General"
    Range("I2").FormulaR1C1 = "=SUM(RC[-2]/RC[-4])"
''  Copy down
    Range("I2:I" & TotalRows).FillDown
End Sub
