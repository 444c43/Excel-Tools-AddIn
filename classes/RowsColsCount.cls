VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RowsColsCount"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private row_count%
Private column_count%
Private sheet_name$

Public Sub GetRowsAndCols(sheet_name$)
    Sheets(sheet_name).Select
    row_count = Sheets(sheet_name).Range("A65536").End(xlUp).Row
    column_count = Sheets(sheet_name).Range("IV1").End(xlToLeft).Column
End Sub

'Properties
Public Property Get SheetName$()
    SheetName = sheet_name
End Property
Public Property Let SheetName(Value$)
    sheet_name = Value
End Property

Public Property Get RowCount%()
    RowCount = row_count
End Property

Public Property Get ColumnCount%()
    ColumnCount = column_count
End Property





