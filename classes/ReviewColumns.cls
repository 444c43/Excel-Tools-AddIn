VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ReviewColumns"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Sub CreateReviewColumns(review_weeks#, review_formula$)
    Dim last_row&
    last_row = Sheets("Serial File").Range("A65535").End(xlUp).Row
    
    Call RemoveColumnErrors(last_row)
    Call AddRemoveColumns
    Call AddHeaders
    Call AddFormulas(review_weeks, review_formula)
    Call FillDownAndPaste(last_row)
    Call CleanUp(last_row)
    Cells.EntireColumn.AutoFit
    Call AdjustColumnWidths
    Call FormatFirstRow
    Call TextToColumns
    Call AddNumberFormat(last_row)
    Call AddAvgScans(last_row, review_weeks)
End Sub
Private Sub RemoveColumnErrors(last_row&)
    Call RemoveNumberAsTextError("C2:C" & last_row)
    Call RemoveNumberAsTextError("E2:E" & last_row)
    Call RemoveNumberAsTextError("I2:I" & last_row)
    Call RemoveNumberAsTextError("M2:M" & last_row)
    Call RemoveNumberAsTextError("O2:O" & last_row)
    Call RemoveNumberAsTextError("P2:P" & last_row)
End Sub
Private Sub RemoveNumberAsTextError(cellRange As String)
    Dim uniqueCell As Range
    
    For Each uniqueCell In Range(cellRange)
        With uniqueCell.Errors(xlNumberAsText)
            If .Value = True Then
                .Ignore = True
            End If
        End With
    Next uniqueCell
End Sub
Private Sub AddRemoveColumns()
    Sheets("Serial File").Range("B:C,L:N,S:W").Delete Shift:=xlToLeft
    Sheets("Serial File").Columns("J:J").Cut
    Sheets("Serial File").Range("F:F").Insert Shift:=xlToRight
    Sheets("Serial File").Range("F:G").Insert Shift:=xlToRight
    Sheets("Serial File").Range("I:J").Insert Shift:=xlToRight
End Sub
Private Sub AddHeaders()
    'headers new columns
    With Sheets("Serial File")
        .Range("F1").Value = "Wkly Avg"
        .Range("G1").Value = "Ship Qty"
        .Range("I1").Value = "Proposed Bin Sys"
        .Range("J1").Value = "Proposed Add/Rem"
        .Range("R1").Value = "Pc Price"
        .Range("S1").Value = "MLV"
        .Range("T1").Value = "Net Chg Value"
        .Range("U1").Value = "PLV"
        .Range("V1").Value = "Serial Status"
        .Range("W1").Value = "To Review"
        .Range("X1").Value = "Sales < $15"
        .Range("Y1").Value = "Missing Pc Price"
        .Range("F:G,I:J,R:U").NumberFormat = "General"
    End With
End Sub
Private Sub AddFormulas(review_weeks#, review_formula$)
    Sheets("Review Data").Columns("A:A").TextToColumns Destination:=Range("A1"), DataType:=xlDelimited
    
    With Sheets("Serial File")
        .Range("F2").FormulaR1C1 = "=ROUNDUP((RC[1]/" & review_weeks & "),0)"
        .Range("G2").FormulaR1C1 = "=IF(ISERROR(VLOOKUP(RC[-6],'Review Data'!R2C1:R" & Sheets("Review Data").Range("A36000").End(xlUp).Row & "C24,24,False))=TRUE,0,VLOOKUP(RC[-6],'Review Data'!R2C1:R" & Sheets("Review Data").Range("A36000").End(xlUp).Row & "C24,24,False))"
        .Range("I2").FormulaR1C1 = "=SUM(RC[1],RC[-1])" 'SUM(L2,I2)
        .Range("J2").Formula = review_formula
        .Range("R2").FormulaR1C1 = "=VLOOKUP(RC[-14],'Price List'!R2C3:R" & Sheets("Price List").Range("A36000").End(xlUp).Row & "C6,4,False)"
        .Range("S2").FormulaR1C1 = "=SUM((RC[-1]*RC[-14])*RC[-11])"
        .Range("T2").FormulaR1C1 = "=SUM((RC[-2]*RC[-15])*RC[-10])"
        .Range("U2").FormulaR1C1 = "=SUM((RC[-3]*RC[-16])*RC[-12])"
    End With
End Sub

Private Sub FillDownAndPaste(last_row&)
    With Sheets("Serial File")
        .Range("F2:G" & last_row & ",I2:J" & last_row & ",R2:U" & last_row).FillDown
        .Columns("G:G").Copy
        .Range("G1").PasteSpecial Paste:=xlPasteValues
        .Columns("R:R").Copy
        .Range("R1").PasteSpecial Paste:=xlPasteValues
    End With
End Sub

Private Sub CleanUp(last_row&)
    Dim i%
    For i = 2 To last_row
        If Sheets("Serial File").Range("R" & i).Text = "#N/A" Then
            Sheets("Serial File").Range("R" & i).Value = 0
            Sheets("Serial File").Range("A" & i).Interior.ColorIndex = 3
            Sheets("Serial File").Range("Y" & i).Value = "x"
        End If
    Next i
End Sub

Private Sub AdjustColumnWidths()
    With Sheets("Serial File")
        .Range("C:D,V:V").ColumnWidth = 15
        .Range("F:F,R:R,X:X").ColumnWidth = 11
        .Range("G:G,Y:Y").ColumnWidth = 8
        .Columns("H:H").ColumnWidth = 7
        .Range("I:I,W:W").ColumnWidth = 10
        .Columns("J:J").ColumnWidth = 12
    End With
End Sub

Private Sub FormatFirstRow()
    With Sheets("Serial File")
        .Range("A1:Y1").HorizontalAlignment = xlCenter
        .Range("A1:Y1").VerticalAlignment = xlCenter
        .Range("A1:Y1").WrapText = True
        .Range("A1:Y1").Font.Bold = True
        .Range("I1:J1").Interior.ColorIndex = 4
    End With
End Sub

Private Sub TextToColumns()
    With Sheets("Serial File")
        .Columns("C:C").TextToColumns Destination:=Range("C1"), DataType:=xlDelimited
        .Columns("H:H").TextToColumns Destination:=Range("H1"), DataType:=xlDelimited
        .Columns("L:L").TextToColumns Destination:=Range("L1"), DataType:=xlDelimited
        .Columns("O:O").TextToColumns Destination:=Range("O1"), DataType:=xlDelimited
        
        .Range("A:D,K:M,O:P,V:Y").HorizontalAlignment = xlCenter
    End With
End Sub

Private Sub AddNumberFormat(last_row&)
    With Sheets("Serial File")
        .Range("F2:G" & last_row).NumberFormat = "#,##0"
        .Range("F2:G" & last_row).HorizontalAlignment = xlRight
        
        .Range("R2:R" & last_row).NumberFormat = "$#,##0.00000"
        .Range("R2:R" & last_row).HorizontalAlignment = xlRight
        
        .Range("S2:U" & last_row).NumberFormat = "$#,##0.00"
        .Range("S2:U" & last_row).HorizontalAlignment = xlRight
    End With
End Sub

Private Sub AddAvgScans(last_row&, review_weeks#)
    With Sheets("Serial File")
        .Columns("I:I").Insert Shift:=xlToRight
        .Range("I1").Value = "Total Bin Scans"
        .Columns.NumberFormat = "General"
        .Range("I2").FormulaR1C1 = "=ROUNDUP(SUM(RC[-2]/RC[-4]),0)"
        .Range("I2:I" & last_row).FillDown
    End With
End Sub
