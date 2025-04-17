Sub GEN_ROW()
Dim i As Integer, j As Integer, s As String, r As Integer
r = 1
For i = 1 To 30
    s = "F" + Str(i)
    For j = 1 To 30
        r = r + 1
        Cells(1, r) = s
    Next j
Next i
End Sub

Sub GEN_COL()
Dim i As Integer, j As Integer, s As String, r As Integer
r = 1
For i = 1 To 30
    For j = 1 To 30
        s = "F" + Str(j)
        r = r + 1
        Cells(2, r) = s
    Next j
Next i
End Sub

Sub Fill_Right()
    Selection.AutoFill Destination:=Range("B4:AHQ4"), Type:=xlFillDefault
    Range("B4:V4").Select
    Range("B4").Select
End Sub

Sub Copy_sheet()
Dim i As Integer, s As String, prevs As String

For i = 2 To 50
    s = "Ans" + Str(i)
    s = WorksheetFunction.Trim(s)
    prevs = "Ans" + Str(i - 1)
    prevs = WorksheetFunction.Trim(prevs)
    
    Sheets("Ans 1").Select
    Sheets("Ans 1").Copy After:=Sheets(prevs)
    Sheets("Ans 1 (2)").Select
    Sheets("Ans 1 (2)").Name = s
    
    Range("B6").Select
    ActiveCell.FormulaR1C1 = i

Next i
End Sub
