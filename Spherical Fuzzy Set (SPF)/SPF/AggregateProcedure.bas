Sub SPFSWAMAGG()
Dim i As Integer, j As Integer, k As Integer
Dim x As Range, y As Range

num_row = 10
num_col = 16
num_exp = 3

exp_row_offset = 13
exp_col_offset = 1

agg_row_offset = 1
agg_col_offset = 2

exp_code = "Ans"

For i = 1 To num_row
    For j = 1 To num_col
        For k = 1 To num_exp
            s = exp_code & k
            Cells(1, k + 1) = Sheets(s).Cells(exp_row_offset + i, exp_col_offset + j)
        Next k
        If Cells(1, 1) <> 0 Then
        Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset) = "=SPFSWAM(B1:D1,B2:D2)"
        Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset).Copy
        Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End If
        If Cells(1, 2) = 0 Then Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset) = 0
        Next j
Next i
End Sub

Sub SPFSWGMAGG()
Dim i As Integer, j As Integer, k As Integer
Dim x As Range, y As Range

num_row = 3
num_col = 3
num_exp = 3

exp_row_offset = 1
exp_col_offset = 1

agg_row_offset = 0
agg_col_offset = 0

exp_code = "Ans"

For i = 1 To num_row
    For j = 1 To num_col
        For k = 1 To num_exp
            s = exp_code & k
            Cells(1, k + 1) = Sheets(s).Cells(exp_row_offset + i, exp_col_offset + j)
        Next k
        If Cells(1, 1) <> 0 Then
        Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset) = "=SPFSWGM(B1:D1,B2:D2)"
        Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset).Copy
        Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End If
        If Cells(1, 2) = 0 Then Cells(i + 3 + agg_row_offset, j + 1 + agg_col_offset) = 0
        Next j
Next i
End Sub