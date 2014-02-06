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

Private acct_name$
Private acct_number$
Private acct_type$
Private del_frequency$
Private review_period$
Private acct_formula$
Private whse$
Private acct_pricecode$
Private review_weeks#

Private CustomerList As CustomerData
Private CollectionIndex%

Public Sub SetupCustomerData(number$)
    Set CustomerList = New CustomerData
    CustomerList.Setup
    acct_number = number
    Call FindCollectionIndex
    Call SetProperties
End Sub

Private Sub FindCollectionIndex()
    Dim x%
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
Public Property Get name$()
    name = acct_name
End Property

Public Property Get AcctNumber$()
    AcctNumber = acct_number
End Property

Public Property Get AcctType$()
    AcctType = acct_type
End Property

Public Property Get DeliveryFrequency$()
    DeliveryFrequency = del_frequency
End Property

Public Property Get ShippingWHSE$()
    ShippingWHSE = whse
End Property

Public Property Get pricecode$()
    pricecode = acct_pricecode
End Property

'PUBLICLY SET PROPERTIES
Public Property Get ReviewPeriod$()
    ReviewPeriod = review_period
End Property
Public Property Let ReviewPeriod(Value$)
    review_period = Value
End Property

Public Property Get ReviewWeeks#()
    ReviewWeeks = review_weeks
End Property
Public Property Let ReviewWeeks(Value#)
    review_weeks = Value
End Property

Public Property Get Formula$()
    Formula = acct_formula
End Property
Public Property Let Formula(Value$)
    acct_formula = Value
End Property
