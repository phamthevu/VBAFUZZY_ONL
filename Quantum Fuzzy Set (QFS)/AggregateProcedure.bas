Sub NQFSAGGWDEMATEL()
Dim i As Integer, j As Integer, k As Integer
Dim x As Range, y As Range
For i = 1 To 20
    For j = 1 To 20
        For k = 1 To 58
            s = "Ans " & k
            Cells(1, k) = Sheets(s).Cells(27 + i, 1 + j)
        Next k
            If Cells(1, 1) <> 0 Then
            Cells(i + 4, j + 1) = "=QFSAGGW(A1:BF1,A2:BF2)"
            Cells(i + 4, j + 1).Copy
            Cells(i + 4, j + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End If
        If Cells(1, 1) = 0 Then Cells(i + 4, j + 1) = 0
        Next j
Next i
End Sub


Sub NQFSAGGWTOPSIS()
Dim i As Integer, j As Integer, k As Integer
Dim x As Range, y As Range
For i = 1 To 7
    For j = 1 To 20
        For k = 1 To 58
            s = "Ans " & k
            Cells(1, k) = Sheets(s).Cells(14 + i, 1 + j)
        Next k
        If Cells(1, 1) <> 0 Then
        Cells(i + 4, j + 1) = "=QFSAGGW(A1:BF1,A2:BF2)"
        Cells(i + 4, j + 1).Copy
        Cells(i + 4, j + 1).PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
        End If
        If Cells(1, 1) = 0 Then Cells(i + 4, j + 1) = 0
        Next j
Next i
End Sub