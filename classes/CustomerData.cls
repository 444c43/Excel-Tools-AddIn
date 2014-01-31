VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CustomerData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public CustomerNames As Collection
Public CustomerNumbers As Collection
Public CustomerDeliveryFrequency As Collection
Public ShippingWHSE As Collection
Public ProgramType As Collection
Public PriceCode As Collection

Public Sub Setup()
    Call Instantiate
    Call SetupNames
    Call SetupNumbers
    Call SetupCustomerDeliveryFrequency
    Call SetupShippingWHSE
    Call SetupProgramType
    Call SetupPriceCode
End Sub

Private Sub Instantiate()
    Set CustomerNames = New Collection
    Set CustomerNumbers = New Collection
    Set CustomerDeliveryFrequency = New Collection
    Set ShippingWHSE = New Collection
    Set ProgramType = New Collection
    Set PriceCode = New Collection
End Sub

''''''''''''''''''''''''''''''''''''''''
'change values below to maintain
'customer data as needed
''''''''''''''''''''''''''''''''''''''''
Private Sub SetupNames()
    CustomerNames.Add "AM General"
    CustomerNames.Add "Atlas Copco - Garland"
    CustomerNames.Add "Atlas Copco - McKinney"
    CustomerNames.Add "Atlas Copco - Steco"
    CustomerNames.Add "Carrier"
    CustomerNames.Add "Cequent"
    CustomerNames.Add "Club Car"
    CustomerNames.Add "Dana Lugoff"
    CustomerNames.Add "Eaton - Electric"
    CustomerNames.Add "Eaton - Kings Mt."
    CustomerNames.Add "Eaton - Shen"
    CustomerNames.Add "Eaton - SLP"
    CustomerNames.Add "JLG - Bedford"
    CustomerNames.Add "JLG - McConnellsburg"
    CustomerNames.Add "JLG - Military"
    CustomerNames.Add "JLG - Orrville"
    CustomerNames.Add "JLG - Service Plus"
    CustomerNames.Add "JLG - Shippensburg"
    CustomerNames.Add "Meritor - Fletcher"
    CustomerNames.Add "Meritor - Forest City"
    CustomerNames.Add "Meritor - Heath"
    CustomerNames.Add "Meritor - Maxton"
    CustomerNames.Add "Square D Seneca"
    CustomerNames.Add "Stryker"
End Sub

Private Sub SetupNumbers()
    CustomerNumbers.Add "0012973"
    CustomerNumbers.Add "0051530"
    CustomerNumbers.Add "0051532"
    CustomerNumbers.Add "0051533"
    CustomerNumbers.Add "0024116"
    CustomerNumbers.Add "0075351"
    CustomerNumbers.Add "0028740"
    CustomerNumbers.Add "0031818"
    CustomerNumbers.Add "0037550"
    CustomerNumbers.Add "0037450"
    CustomerNumbers.Add "0037600"
    CustomerNumbers.Add "0037420"
    CustomerNumbers.Add "0054061"
    CustomerNumbers.Add "0054060"
    CustomerNumbers.Add "0054067"
    CustomerNumbers.Add "0054062"
    CustomerNumbers.Add "0054065"
    CustomerNumbers.Add "0054063"
    CustomerNumbers.Add "0076269"
    CustomerNumbers.Add "0076259"
    CustomerNumbers.Add "0076270"
    CustomerNumbers.Add "0076266"
    CustomerNumbers.Add "0081958"
    CustomerNumbers.Add "0084055"
End Sub

Private Sub SetupCustomerDeliveryFrequency()
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Bi-Wkly"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "4x/Wk"
    CustomerDeliveryFrequency.Add "4x/Wk"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Daily"
    CustomerDeliveryFrequency.Add "Weekly"
    CustomerDeliveryFrequency.Add "Daily"
End Sub
Private Sub SetupShippingWHSE()
    ShippingWHSE.Add "26"
    ShippingWHSE.Add "25"
    ShippingWHSE.Add "25"
    ShippingWHSE.Add "25"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "26"
    ShippingWHSE.Add "6"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "18"
    ShippingWHSE.Add "75"
    ShippingWHSE.Add "77"
    ShippingWHSE.Add "77"
    ShippingWHSE.Add "77"
    ShippingWHSE.Add "77"
    ShippingWHSE.Add "77"
    ShippingWHSE.Add "77"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "15"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "2"
    ShippingWHSE.Add "44"
End Sub
Private Sub SetupProgramType()
    ProgramType.Add "1 Wk"
    ProgramType.Add "3 Wk"
    ProgramType.Add "3 Wk"
    ProgramType.Add "3 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "2 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "5 Day"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "1 Wk"
    ProgramType.Add "3 Wk"
    ProgramType.Add "1 Wk"
End Sub

Private Sub SetupPriceCode()
    PriceCode.Add "014"
    PriceCode.Add "635"
    PriceCode.Add "635"
    PriceCode.Add "635"
    PriceCode.Add "130"
    PriceCode.Add "683"
    PriceCode.Add "042"
    'DANA LUGOFF INCORRECT
    PriceCode.Add "000"
    PriceCode.Add "438"
    PriceCode.Add "438"
    PriceCode.Add "438"
    PriceCode.Add "438"
    PriceCode.Add "364"
    PriceCode.Add "364"
    PriceCode.Add "364"
    PriceCode.Add "364"
    PriceCode.Add "364"
    PriceCode.Add "364"
    PriceCode.Add "189"
    PriceCode.Add "189"
    PriceCode.Add "189"
    PriceCode.Add "189"
    PriceCode.Add "889"
    PriceCode.Add "478"
End Sub

