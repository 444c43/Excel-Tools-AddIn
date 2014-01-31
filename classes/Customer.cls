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

Private acct_name As String
Private acct_number As String
Private acct_type As String
Private del_frequency As String
Private review_period As String
Private acct_formula As String
Private whse As String
Private acct_pricecode As String
Private review_weeks As Double

Private CustomerList As CustomerData
Private CollectionIndex As Integer

Public Sub SetupCustomerData(number As String)
    Set CustomerList = New CustomerData
    CustomerList.Setup
    acct_number = number
    Call FindCollectionIndex
    Call SetProperties
End Sub

Private Sub FindCollectionIndex()
    Dim x As Integer
    For x = 1 To CustomerList.CustomerNumbers.count
        If CustomerList.CustomerNumbers(x) = acct_number Then
            CollectionIndex = x
        End If
    Next x
End Sub

Private Sub SetProperties()
    acct_name = CustomerList.CustomerNames(CollectionIndex)
    acct_type = CustomerList.ProgramType(CollectionIndex)
    whse = CustomerList.ShippingWHSE(CollectionIndex)
    del_frequency = CustomerList.CustomerDeliveryFrequency(CollectionIndex)
    acct_pricecode = CustomerList.pricecode(CollectionIndex)
End Sub

'PRIVATELY SET PROPERTIES
Public Property Get name() As String
    name = acct_name
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

Public Property Get pricecode() As String
    pricecode = acct_pricecode
End Property

'PUBLICLY SET PROPERTIES
Public Property Get ReviewPeriod() As String
    ReviewPeriod = review_period
End Property
Public Property Let ReviewPeriod(Value As String)
    review_period = Value
End Property

Public Property Get ReviewWeeks() As Double
    ReviewWeeks = review_weeks
End Property
Public Property Let ReviewWeeks(Value As Double)
    review_weeks = Value
End Property

Public Property Get Formula() As String
    Formula = acct_formula
End Property
Public Property Let Formula(Value As String)
    acct_formula = Value
End Property
