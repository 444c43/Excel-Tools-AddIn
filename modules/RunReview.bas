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
    Call Constructor
    Call InitializeObjects
    
    If CountErrors = 0 Then
        Call Destructors
        Call ExecuteProgramReview
    Else
        Call UserMessage.DisplayErrors(ReviewValidations.ValidationErrors)
    End If
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
    Call AdjustAllSheetHeaders
    Call InactiveSerialsTab
    Call PartsNotOrderedTab
End Sub

'ALL SUBS AND FUNCTIONS BELOW ARE CALLED FROM ExecuteProgramReview
Private Sub AdjustAllSheetHeaders()
    Dim Headers As AS400Headers
    Set Headers = New AS400Headers
    
    Call Headers.Convert("Serial File")
    Call Headers.Convert("Review Data")
    Call Headers.Convert("Price List")
End Sub
Private Sub InactiveSerialsTab()
    Dim InactiveTab As InactiveSerials
    Set InactiveTab = New InactiveSerials
    
    Sheets.Add After:=Sheets("Price List")
    activesheet.name = "Inactive Serials"
    Call InactiveTab.CutDeletedCopyInactive
End Sub

Private Sub PartsNotOrderedTab()
    Dim UnorderedParts As NotOrdered
    Set UnorderedParts = New NotOrdered
    
    Sheets.Add After:=Sheets("Inactive Serials")
    activesheet.name = "Parts Not Ordered"
    Call UnorderedParts.AddHeaders
    Call UnorderedParts.FindNotOrderedParts
    Call UnorderedParts.AddGFCPartNumbers
End Sub

