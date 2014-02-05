Attribute VB_Name = "RunReview"
Option Explicit
Public ReviewValidations As New Validations
Public UserMessage As New Message
Public ReviewCustomer As Customer

Sub SetupSheets()
    Dim ReviewSheets As SheetSetups
    Set ReviewSheets = New SheetSetups
    
    On Error GoTo DisplayError
    Call ReviewSheets.AdjustSheets(Array("Serial File", "Review Data", "Price List"))
    Sheets("Serials").Select
    Exit Sub

DisplayError:
    MsgBox "Error: A sheet named Serials, Review or Price already exists!"
    activesheet.Delete
    Exit Sub
End Sub

Sub EntryPoint()
    Application.ScreenUpdating = False
    Call Constructor
    Call InitializeObjects
    
    If CountErrors = 0 Then
        Call Destructors
        Call ExecuteProgramReview
    Else
        Call UserMessage.DisplayErrors(ReviewValidations.ValidationErrors)
    End If
    Application.ScreenUpdating = True
End Sub
Private Sub Constructor()
    Set ReviewValidations = New Validations
    Set UserMessage = New Message
    Set ReviewCustomer = New Customer
End Sub
Private Sub InitializeObjects()
    Call ReviewValidations.SetupObject("LIST80", "GFCCS1")
    Call ReviewCustomer.SetupCustomerData("GFCCS1")
End Sub
Private Function CountErrors%()
    Call ReviewValidations.ValidateSheetNames(Array("Serial File", "Review Data", "Price List"))
    Call ReviewValidations.ValidateHeaders(Array("GFCSR#", "SERIAL", "CONO80"))
    Call ReviewValidations.ValidateCustomerData(ReviewCustomer.AcctNumber, ReviewCustomer.PriceCode)
    CountErrors = ReviewValidations.ValidationErrors.count
End Function
Private Sub Destructors()
    Set ReviewValidations = Nothing
    Set UserMessage = Nothing
End Sub

' ACTUAL REVIEW EXECUTES HERE
Private Sub ExecuteProgramReview()
    frmReview.Show
    'EXECUTE PROGRAM HERE ALL VALIDATIONS PASS
    Call RemoveZeroShipAndZeroBinQty
    Call AdjustAllSheetHeaders
    Call NotScannedSerialsTab
    Call InactiveSerialsTab
    Call PartsNotOrderedTab
    Call MissingPcPriceTab
    Call AddReviewColumns
    Call CalculateSerialStatus
    Call RemoveSheets
    Call FormatSheets
End Sub

Private Sub RemoveZeroShipAndZeroBinQty()
    Dim i&, last_row&
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    For i = last_row To 2 Step -1
        If Sheets("Serial File").Range("O" & i).value = 0 And _
            Application.WorksheetFunction.CountIf(Sheets("Review Data").Range("A:A"), Sheets("Serial File").Range("A" & i)) = 0 Then
                Sheets("Serial File").Range("A" & i).EntireRow.Delete
        End If
    Next i
End Sub

'ALL SUBS AND FUNCTIONS BELOW ARE CALLED FROM ExecuteProgramReview
Private Sub AdjustAllSheetHeaders()
    Dim Headers As AS400Headers
    Set Headers = New AS400Headers
    
    Call Headers.Convert("Serial File")
    Call Headers.Convert("Review Data")
    Call Headers.Convert("Price List")
End Sub

Private Sub NotScannedSerialsTab()
    Dim NotScannedTab As NotScannedSerials
    Set NotScannedTab = New NotScannedSerials
    
    Sheets.Add After:=Sheets("Price List")
    activesheet.name = "Not Scanned"
    
    Call NotScannedTab.SetupTab
End Sub

Private Sub InactiveSerialsTab()
    Dim InactiveTab As InactiveSerials
    Set InactiveTab = New InactiveSerials
    
    Sheets.Add After:=Sheets("Not Scanned")
    activesheet.name = "Inactive Serials"
    
    Call InactiveTab.CutDeletedCopyInactive
End Sub

Private Sub PartsNotOrderedTab()
    Dim UnorderedParts As NotOrdered
    Set UnorderedParts = New NotOrdered
    
    Sheets.Add After:=Sheets("Inactive Serials")
    activesheet.name = "Parts Not Ordered"
    
    UnorderedParts.SetupNotOrdered
End Sub

Private Sub MissingPcPriceTab()
    Dim MissingPcPrice As MissingPrice
    Set MissingPcPrice = New MissingPrice
    
    Sheets.Add After:=Sheets("Parts Not Ordered")
    activesheet.name = "Missing Pc Price"
    
    MissingPcPrice.SetupMissingPcPrice
End Sub

Private Sub AddReviewColumns()
    Dim NewColumns As ReviewColumns
    Set NewColumns = New ReviewColumns
    
    Call NewColumns.CreateReviewColumns(ReviewCustomer.ReviewWeeks, ReviewCustomer.Formula)
End Sub

Private Sub CalculateSerialStatus()
    Dim SerialStatus As SerialCalculations
    Set SerialStatus = New SerialCalculations
   
    SerialStatus.CalculateSerials
    SerialStatus.MoveSerialStatus
End Sub

Private Sub RemoveSheets()
    Application.DisplayAlerts = False
    Sheets("Price List").Delete
    Sheets("Review Data").Delete
    Application.DisplayAlerts = True
End Sub

Private Sub FormatSheets()
    Dim WorksheetFormat As WorksheetFormatting
    Set WorksheetFormat = New WorksheetFormatting
    
    Call WorksheetFormat.FormatAllWorksheets(ReviewCustomer.name, ReviewCustomer.AcctNumber, ReviewCustomer.ReviewPeriod)
End Sub


Sub test()
Range("I2").Formula = "=IF(ROUNDUP(SUM((G2/13)/E2),0)=1,2,IF(AND(ROUNDUP(SUM((G2/13)/E2),0)=0,H2>2),2,ROUNDUP(SUM((G2/13)/E2),0)))"


End Sub

