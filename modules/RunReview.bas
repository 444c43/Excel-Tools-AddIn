Attribute VB_Name = "RunReview"
Option Explicit
Public ReviewValidations As New Validations
Public UserMessage As New Message
Public ReviewCustomer As Customer
Private NewReviewSheets As ReviewSheets
Private CustomerParts As UniqueContent

Sub SetupSheets()
    Dim ReviewSheets As SheetSetups
    Set ReviewSheets = New SheetSetups
    
    On Error GoTo DisplayError
    Call ReviewSheets.AdjustSheets(Array("Serial File", "Review Data", "Price List"))
    Sheets("Serial File").Select
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
    Set CustomerParts = New UniqueContent
End Sub
Private Sub InitializeObjects()
    Call ReviewValidations.SetupObject("LIST80", "GFCCS1")
    Call ReviewCustomer.SetupReviewCustomer("GFCCS1")
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
    Call AddNewTabs
    Call AddReviewColumns
    Call CalculateSerialStatus
    Call SetupPartsNotOrderedTab
    Call RemoveSheets
    Call FormatSheets
    Sheets(ReviewCustomer.AcctNumber).Select
End Sub

Private Sub RemoveZeroShipAndZeroBinQty()
    Dim i&, last_row&
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    For i = last_row To 2 Step -1
        If Sheets("Serial File").Range("O" & i).Value = 0 And _
            Application.WorksheetFunction.CountIf(Sheets("Review Data").Range("A:A"), Sheets("Serial File").Range("A" & i)) = 0 Then
                Sheets("Serial File").Range("A" & i).EntireRow.Delete
        End If
    Next i
End Sub

'ALL SUBS AND FUNCTIONS BELOW ARE CALLED FROM ExecuteProgramReview
Private Sub AdjustAllSheetHeaders()
    Dim Headers As AS400
    Set Headers = New AS400
    
    Call Headers.ConvertHeaders("Serial File")
    Call Headers.ConvertHeaders("Review Data")
    Call Headers.ConvertHeaders("Price List")
End Sub

Private Sub AddNewTabs()
    Set NewReviewSheets = New ReviewSheets
    Call NewReviewSheets.InstantiateVariables
    
    Call SetupNotScannedTab
    Call SetupInactiveTab
    Call SetupMissingPcPriceTab
End Sub
Private Sub SetupNotScannedTab()
    'NotScannedTab
    Call NewReviewSheets.AddNewSheet("Not Scanned", "Price List")
End Sub
Private Sub SetupInactiveTab()
    'Inactive Sheet
    Call NewReviewSheets.AddNewSheet("Inactive Serials", "Not Scanned")
    Call NewReviewSheets.CopyHeaders("Serial File", "Inactive Serials")
    Call NewReviewSheets.CutDeletedCopyInactive
End Sub
Private Sub SetupMissingPcPriceTab()
    'Missing Piece Price
    Call NewReviewSheets.AddNewSheet("Missing Pc Price", "Inactive Serials")
    Call NewReviewSheets.SetupMissingPcPrice
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
    Call CleanUpNotScannedTab
    SerialStatus.MoveSerialStatus
End Sub

Private Sub SetupPartsNotOrderedTab()
    'create UniqueContent object
    Dim UnorderedParts As UniqueContent
    Set UnorderedParts = New UniqueContent
    
    'set search to Main tab with header name, compile list
    Call UnorderedParts.Initialize("Serial File", "Customer Part #")
    Call UnorderedParts.GetNotOrdered
    
    'reset search to Inactive tab with header name, add to compiled list
    Call UnorderedParts.SetSearchColumn("Inactive Serials", "Customer Part #")
    Call UnorderedParts.GetInactive("Serial File")
    
    'setup new sheet, with compiled list
    Call NewReviewSheets.AddNewSheet("Parts Not Ordered", "Inactive Serials")
    Call NewReviewSheets.SetupPartsNotOrdered(UnorderedParts.UniqueItems)
End Sub

Private Sub CleanUpNotScannedTab()
    Sheets("Not Scanned").Range("A1:AA1").Value = Sheets("Serial File").Range("A1:AA1").Value
    Sheets("Not Scanned").Range("F:G,I:K,S:Z").Delete Shift:=xlToLeft
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
