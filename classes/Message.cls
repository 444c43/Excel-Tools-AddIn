VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Message"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function DisplayErrors(error_list As Collection)
    Dim item
    Debug.Print "The following problems need corrected:"
    For Each item In error_list
        Debug.Print CStr(item)
    Next item
End Function
