Function QFSUC(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
t = Trim(Mid(s, 2, i - 2))
j = WorksheetFunction.Find("e", t, 1)
QFSUC = CDbl(Mid(t, 1, j - 1))
End Function

Function QFSUE(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
t = Trim(Mid(s, 2, i - 2))
j = WorksheetFunction.Find("e", t, 1)
QFSUE = CDbl(Mid(t, j + 5, Len(t) - j))
End Function

Function QFSVC(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, i + 2, j - i - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSVC = CDbl(Mid(t, 1, k - 1))
End Function

Function QFSVE(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, i + 2, j - i - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSVE = CDbl(Mid(t, k + 5, Len(t) - 1))
End Function
Function QFSHC(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, j + 2, Len(s) - j - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSHC = CDbl(Mid(t, 1, k - 1))
End Function

Function QFSHE(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, j + 2, Len(s) - j - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSHE = CDbl(Mid(t, k + 5, Len(t) - 1))
End Function






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

Function SUMQFSC(x As Range, c As Integer) As Double
Dim i As Integer, j As Integer, k As Integer
k = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)

If c = 1 Then
    For i = 1 To k
    If x(i) <> 0 Then SUMQFSC = SUMQFSC + QFSUC(x(i))
    Next i
End If

If c = 2 Then
    For i = 1 To k
    If x(i) <> 0 Then SUMQFSC = SUMQFSC + QFSVC(x(i))
    Next i
End If

If c = 3 Then
    For i = 1 To k
    If x(i) <> 0 Then SUMQFSC = SUMQFSC + QFSHC(x(i))
    Next i
End If
End Function

Function SUMQFSE(x As Range, c As Integer) As Double
Dim i As Integer, j As Integer, k As Integer
k = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)

If c = 1 Then
    For i = 1 To k
    If x(i) <> 0 Then SUMQFSE = SUMQFSE + QFSUE(x(i))
    Next i
End If

If c = 2 Then
    For i = 1 To k
    If x(i) <> 0 Then SUMQFSE = SUMQFSE + QFSVE(x(i))
    Next i
End If

If c = 3 Then
    For i = 1 To k
    If x(i) <> 0 Then SUMQFSE = SUMQFSE + QFSHE(x(i))
    Next i
End If
End Function



