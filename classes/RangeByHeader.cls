VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RangeByHeader"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Public LastRow%, LastColumn%

Public HeaderRange As Range
Public ColumnRange As Range

Public Sub SetProperties(SheetRef$, HeaderVal$)
    Sheets(SheetRef).Select
    LastRow = ActiveCell.SpecialCells(xlLastCell).Row
    LastColumn = Range("IV1").End(xlToLeft).Column
    
    Set HeaderRange = Range(Cells(1, 1), Cells(1, LastColumn))
    Set ColumnRange = Range(HeaderRange.Find(HeaderVal).Address)
End Sub

