VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Customer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'privately set properties
Private acct_number$
Private acct_name$, acct_type$, whse$
Private del_frequency$, acct_pricecode$
'publicly set properties
Private review_period$, acct_formula$
Private review_weeks#
'row and column variables
Private LastColumn%, LastRow%
'objects
Private CustomerList As CustomerData
Private SearchRange, CustomerNumber As Range

Public Sub SetupCustomerData(header As String)
    LastRow = ActiveCell.SpecialCells(xlLastCell).Row
    LastColumn = Range("IV1").End(xlToLeft).Column
    Call SetObjects(header)
    CustomerList.Setup
    acct_number = CustomerNumber.Cells(2, 1)
    Call SetProperties(FindCollectionIndex)
    Call Destructor
End Sub
Private Sub SetObjects(header_text$)
    Set CustomerList = New CustomerData
    Set SearchRange = Range(Cells(1, 1), Cells(1, LastColumn))
    Set CustomerNumber = Range(SearchRange.Find(header_text).Address)
End Sub

Private Function FindCollectionIndex%()
    Dim i%
    For i = 1 To CustomerList.CustomerNumbers.count
        If CustomerList.CustomerNumbers(i) = acct_number Then
            FindCollectionIndex = i
            Exit Function
        End If
    Next i
End Function

Private Sub SetProperties(index_value%)
    acct_name = CustomerList.CustomerNames(index_value)
    acct_type = CustomerList.ProgramType(index_value)
    whse = CustomerList.ShippingWHSE(index_value)
    del_frequency = CustomerList.CustomerDeliveryFrequency(index_value)
    acct_pricecode = CustomerList.PriceCode(index_value)
End Sub

Private Sub Destructor()
    Set CustomerList = Nothing
    Set SearchRange = Nothing
    Set CustomerNumber = Nothing
End Sub

'PRIVATELY SET PROPERTIES
Public Property Get Name() As String
    Name = acct_name
End Property

Public Property Get AcctNumber() As String
    AcctNumber = acct_number
End Property

Public Property Get AcctType() As String
    AcctType = acct_type
End Property

Public Property Get DeliveryFrequency() As String
    DeliveryFrequency = del_frequency
End Property

Public Property Get ShippingWHSE() As String
    ShippingWHSE = whse
End Property

Public Property Get PriceCode() As String
    PriceCode = acct_pricecode
End Property

'PUBLICLY SET PROPERTIES
Public Property Get ReviewPeriod() As String
    ReviewPeriod = review_period
End Property
Public Property Let ReviewPeriod(value As String)
    review_period = value
End Property

Public Property Get ReviewWeeks() As Double
    ReviewWeeks = review_weeks
End Property
Public Property Let ReviewWeeks(value As Double)
    review_weeks = value
End Property

Public Property Get Formula() As String
    Formula = acct_formula
End Property
Public Property Let Formula(value As String)
    acct_formula = value
End Property
