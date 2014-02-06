VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "DataValidations"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public ErrorList As Collection
Private iterate As Byte

Public Sub ValidateData(customernumber$, pricecode$)
    Set ErrorList = New Collection
    
    If ValidateReviewData(customernumber) = False Then
        ErrorList.Add ("Review Data does not match Serial File data!")
    End If

    If ValidatePriceList(pricecode) = False Then
        ErrorList.Add ("Price List does not match Serial File customer!")
    End If
End Sub

Private Function ValidateReviewData(number$) As Boolean
    ValidateReviewData = (number = sheets("Review Data").Range("B2").Value)
End Function

Private Function ValidatePriceList(pricecode$) As Boolean
    ValidatePriceList = (pricecode = sheets("Price List").Range("B2").Value)
End Function
