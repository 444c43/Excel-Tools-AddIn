VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReview 
   Caption         =   "Review 8.0"
   ClientHeight    =   4605
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5955
   OleObjectBlob   =   "frmReview.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReview"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private x%
Private TotalWeeks#
''============================================================
''  INITIALIZE FORMS
''============================================================
Private Sub UserForm_Initialize()
    Call SetInitialCustomerData
    Call ClearComboBoxes
    Call SetQuarters
    Call SetMonths
    Call SetYears
    Call SetFirstValues
    Call ShowMoWks
''  Set focus to first page
    multiReview.value = 0
End Sub
Private Sub SetInitialCustomerData()
    lblAcctName2.Caption = ReviewCustomer.Name
    lblAcctNumber2.Caption = ReviewCustomer.AcctNumber
    lblAcctType2.Caption = ReviewCustomer.AcctType
End Sub
Private Sub ClearComboBoxes()
    Dim form_object As Control
    
    For Each form_object In Me.Controls
        If TypeOf form_object Is MSForms.ComboBox Then
            form_object.Clear
        End If
    Next form_object
End Sub
Private Sub SetQuarters()
    cbxQuarter.AddItem "Q1 - Jan - Mar"
    cbxQuarter.AddItem "Q2 - Apr - Jun"
    cbxQuarter.AddItem "Q3 - Jul - Sep"
    cbxQuarter.AddItem "Q4 - Oct - Dec"
End Sub
Private Sub SetMonths()
    ' add month ranges
    For x = 1 To 12
        cbxMonth.AddItem left(MonthName(x), 3)
        cbxVariedMo1.AddItem left(MonthName(x), 3)
        cbxVariedMo2.AddItem left(MonthName(x), 3)
    Next x
End Sub
Private Sub SetYears()
    ' add current year plus last three
    For x = Year(Now) To Year(Now) - 3 Step -1
        cbxYear.AddItem x
        cbxQuarterlyYr.AddItem x
        cbxVariedYr1.AddItem x
        cbxVariedYr2.AddItem x
    Next x
End Sub
Private Sub SetFirstValues()
    cbxQuarter.Text = SelectMonth
    cbxQuarterlyYr.Text = cbxQuarterlyYr.list(0)
    cbxMonth.Text = cbxMonth.list(0)
    cbxYear.Text = cbxYear.list(0)
    cbxVariedYr1.Text = cbxVariedYr1.list(0)
    cbxVariedYr2.Text = cbxVariedYr2.list(0)
    cbxVariedMo1.Text = cbxVariedMo1.list(0)
    cbxVariedMo2.Text = cbxVariedMo2.list(11)
End Sub
Private Function SelectMonth$()
    SelectMonth = cbxQuarter.list(2)
    Select Case Month(Now)
        Case 1 To 3
            SelectMonth = cbxQuarter.list(3)
        Case 4 To 6
            SelectMonth = cbxQuarter.list(0)
        Case 7 To 9
            SelectMonth = cbxQuarter.list(1)
    End Select
End Function

''============================================================
''  RUN REVIEW
''============================================================
Private Sub cmdRunReview_Click()
    'determine period range and total weeks to calculate
    If DateChk = True Then
        Call EvaluatePeriods
        Call EvaluateAcctType
        frmReview.Hide
    Else
        MsgBox "Cannot use that date range!"
        SetFirstValues
    End If
End Sub
Private Function DateChk() As Boolean
    If CalculateMonths <= 0 Then
        DateChk = False
    Else
        DateChk = True
    End If
End Function
Private Sub EvaluatePeriods()
    Select Case multiReview.value
        Case 0
            Call SetPeriodAndWeeks((right(cbxQuarter.value, 9) & " " & cbxQuarterlyYr.value), 13)
        Case 1
            Call SetPeriodAndWeeks(cbxMonth.value & " " & cbxYear.value, 4.3)
        Case 2
            Call SetPeriodAndWeeks(cbxVariedMo1.value & " " & cbxVariedYr1.value & _
            " " & cbxVariedMo2.value & " " & cbxVariedYr2.value, TotalWeeks)
    End Select
End Sub
Private Sub SetPeriodAndWeeks(period$, weeks#)
    ReviewCustomer.ReviewPeriod = period
    ReviewCustomer.ReviewWeeks = weeks
End Sub
Private Sub EvaluateAcctType()
    Select Case ReviewCustomer.AcctType
        Case "1 Wk"
            ReviewCustomer.Formula = "=IF(ROUNDUP(SUM((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4]),0)=1,2,ROUNDUP(SUM((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4]),0))"
        Case "2 Wk"
            ReviewCustomer.Formula = "=IF(ROUNDUP(SUM(((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4])*2),0)=1,2,ROUNDUP(SUM(((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4])*2),0))"
        Case "3 Wk"
            ReviewCustomer.Formula = "=IF(ROUNDUP(SUM(((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4])*3),0)=1,2,ROUNDUP(SUM(((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4])*3),0))"
        Case "5 Day"
            ReviewCustomer.Formula = "=IF(ROUNDUP(SUM(((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4])/5),0)=1,2,ROUNDUP(SUM(((RC[-2]/" & ReviewCustomer.ReviewWeeks & ")/RC[-4])/5),0))"
    End Select
End Sub

Private Function CalculateMonths%()
    Dim MonthsYr1 As Date
    Dim MonthsYr2 As Date

    MonthsYr2 = DateSerial(cbxVariedYr2.value, cbxVariedMo2.ListIndex + 1, 1)
    MonthsYr1 = DateSerial(cbxVariedYr1.value, cbxVariedMo1.ListIndex + 1, 1)

    CalculateMonths = DateDiff("m", MonthsYr1, MonthsYr2) + 1
End Function

Private Sub cbxVariedMo1_DropButtonClick()
    Call ShowMoWks
End Sub
Private Sub cbxVariedMo2_DropButtonClick()
    Call ShowMoWks
End Sub

Private Sub cbxVariedYr1_DropButtonClick()
    Call ShowMoWks
End Sub

Private Sub cbxVariedYr2_DropButtonClick()
    Call ShowMoWks
End Sub
Private Sub ShowMoWks()
    Dim TotalMonths%
    TotalMonths = CalculateMonths
    TotalWeeks = Round(CalculateMonths * 4.333, 0)
    lblMo.Caption = "Mo: " & TotalMonths
    lblWks.Caption = "Wks: " & Round(TotalWeeks, 0)
End Sub
Private Sub UserForm_Terminate()
    frmReview.Hide
    Exit Sub
End Sub


