Attribute VB_Name = "RunReview"
Option Explicit
Public ReviewValidations As New Validations
Public UserMessage As New Message

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
    Call ReviewValidations.SetupObject

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
End Sub
Private Function CountErrors%()
    Call ReviewValidations.ValidateSheetNames(Array("Serial File", "Review Data", "Price List"))
    Call ReviewValidations.ValidateHeaders(Array("GFCSR#", "SERIAL", "CONO80"))
    Call ReviewValidations.ValidateCustomerData
    CountErrors = ReviewValidations.ValidationErrors.count
End Function

Private Sub Destructors()
    Set ReviewValidations = Nothing
End Sub
Private Sub ExecuteProgramReview()
    Debug.Print "Success"
End Sub

