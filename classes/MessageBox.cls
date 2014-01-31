VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "MessageBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function MultiLine(boxtype As Variant, title As String, messages As Collection)
    Dim iterate As Byte
    Dim Message As String
    Dim member As Variant
    
    For Each member In messages
        Message = Message + member & vbNewLine
    Next
    
    MsgBox Message, boxtype, title
End Function
