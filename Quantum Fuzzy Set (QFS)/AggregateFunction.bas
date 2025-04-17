Function QFSAGG(x As Range) As String
Dim i As Integer, j As Integer, k As Integer
Dim ucc As Double, uec As Double, tpi As Double, vcc As Double, vec As Double, hcc As Double, hec As Double, a As Double, b As Double, c As Double
k = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)
tpi = 2 * WorksheetFunction.pi()
ucc = (1 - QFSUC(x(1)) ^ 2) ^ (1 / k)
For i = 2 To k
    ucc = ucc * (1 - QFSUC(x(i)) ^ 2) ^ (1 / k)
Next i
ucc = Round((1 - ucc) ^ (1 / 2), 5)

uec = (1 - (QFSUE(x(1)) / tpi) ^ 2) ^ (1 / k)
For i = 2 To k
uec = uec * (1 - (QFSUE(x(i)) / tpi) ^ 2) ^ (1 / k)
Next i
uec = Round((1 - uec) ^ (1 / 2), 5)

vcc = QFSVC(x(1)) ^ (1 / k)
For i = 2 To k
vcc = vcc * QFSVC(x(i)) ^ (1 / k)
Next i
vcc = Round((vcc), 5)

vec = (QFSVE(x(1)) / tpi) ^ (1 / k)
For i = 2 To k
vec = vec * (QFSVC(x(i)) / tpi) ^ (1 / k)
Next i
vec = Round((vec), 5)

a = (1 - QFSUC(x(1)) ^ 2) ^ (1 / k)
b = (1 - QFSUC(x(1)) ^ 2 - QFSHC(x(1)) ^ 2) ^ (1 / k)
For i = 2 To k
a = a * (1 - QFSUC(x(i)) ^ 2) ^ (1 / k)
b = b * (1 - QFSUC(x(i)) ^ 2 - QFSHC(x(i)) ^ 2) ^ (1 / k)
Next i
hcc = Round((a - b) ^ (1 / 2), 5)

a = (1 - (QFSUE(x(1)) / tpi) ^ 2) ^ (1 / k)
b = (1 - (QFSUE(x(1)) / tpi) ^ 2 - (QFSHE(x(1)) / tpi) ^ 2) ^ (1 / k)
For i = 2 To k
a = a * (1 - (QFSUE(x(i)) / tpi) ^ 2) ^ (1 / k)
b = b * (1 - (QFSUE(x(i)) / tpi) ^ 2 - (QFSHE(x(i)) / tpi) ^ 2) ^ (1 / k)
Next i
hec = Round((a - b) ^ (1 / 2), 5)
QFSAGG = "(" & ucc & "ej2p" & uec & "; " & vcc & "ej2p" & vec & "; " & hcc & "ej2p" & hec & ")"
End Function

Function QFSAGGW(x As Range, y As Range) As String
Dim i As Integer, j As Integer, k As Integer
Dim ucc As Double, uec As Double, tpi As Double, vcc As Double, vec As Double, hcc As Double, hec As Double, a As Double, b As Double, c As Double
k = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)
tpi = 2 * WorksheetFunction.pi()
ucc = (1 - QFSUC(x(1)) ^ 2) ^ (y(1))
For i = 2 To k
    ucc = ucc * (1 - QFSUC(x(i)) ^ 2) ^ (y(i))
Next i
ucc = Round((1 - ucc) ^ (1 / 2), 5)

uec = (1 - (QFSUE(x(1)) / tpi) ^ 2) ^ (y(1))
For i = 2 To k
uec = uec * (1 - (QFSUE(x(i)) / tpi) ^ 2) ^ (y(i))
Next i
uec = Round((1 - uec) ^ (1 / 2), 5)

vcc = QFSVC(x(1)) ^ (y(1))
For i = 2 To k
vcc = vcc * QFSVC(x(i)) ^ (y(i))
Next i
vcc = Round((vcc), 5)

vec = (QFSVE(x(1)) / tpi) ^ (y(1))
For i = 2 To k
vec = vec * (QFSVC(x(i)) / tpi) ^ (y(i))
Next i
vec = Round((vec), 5)

a = (1 - QFSUC(x(1)) ^ 2) ^ (y(1))
b = (1 - QFSUC(x(1)) ^ 2 - QFSHC(x(1)) ^ 2) ^ (y(1))
For i = 2 To k
a = a * (1 - QFSUC(x(i)) ^ 2) ^ (y(i))
b = b * (1 - QFSUC(x(i)) ^ 2 - QFSHC(x(i)) ^ 2) ^ (y(i))
Next i
hcc = Round((a - b) ^ (1 / 2), 5)

a = (1 - (QFSUE(x(1)) / tpi) ^ 2) ^ (y(1))
b = (1 - (QFSUE(x(1)) / tpi) ^ 2 - (QFSHE(x(1)) / tpi) ^ 2) ^ (y(1))
For i = 2 To k
a = a * (1 - (QFSUE(x(i)) / tpi) ^ 2) ^ (y(i))
b = b * (1 - (QFSUE(x(i)) / tpi) ^ 2 - (QFSHE(x(i)) / tpi) ^ 2) ^ (y(i))
Next i
hec = Round((a - b) ^ (1 / 2), 5)
QFSAGGW = "(" & ucc & "ej2p" & uec & "; " & vcc & "ej2p" & vec & "; " & hcc & "ej2p" & hec & ")"
End Function

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
