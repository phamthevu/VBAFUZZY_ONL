Function QFSPOW(s As String, k As Double) As String
Dim ucc As Double, uec As Double, tpi As Double, vcc As Double, vec As Double, hcc As Double, hec As Double
Dim i As Integer, pi As Integer
tpi = 2 * WorksheetFunction.pi()
ucc = Round(QFSUC(s) ^ k, 5)
uec = Round((QFSUE(s) / tpi) ^ k, 5)
vcc = Round((1 - (1 - QFSVC(s) ^ 2) ^ k) ^ (1 / 2), 5)
vec = Round((1 - (1 - (QFSVE(s) / tpi) ^ 2) ^ k) ^ (1 / 2), 5)
hhc = Round((((1 - QFSVC(s) ^ 2) ^ k) - ((1 - QFSVC(s) ^ 2 - QFSHC(s) ^ 2) ^ k)) ^ (1 / 2), 5)
hec = Round((((1 - (QFSVE(s) / tpi) ^ 2) ^ k) - ((1 - (QFSHE(s) / tpi) ^ 2 - (QFSVE(s) / tpi) ^ 2) ^ k)) ^ (1 / 2), 5)
QFSPOW = "(" & ucc & "ej2p" & uec & "; " & vcc & "ej2p" & vec & "; " & hhc & "ej2p" & hec & ")"
End Function

Function QFSADD(s1 As String, s2 As String) As String
Dim ucc As Double, uec As Double, tpi As Double, vcc As Double, vec As Double, hcc As Double, hec As Double
Dim i As Integer, pi As Integer
tpi = 2 * WorksheetFunction.pi()
ucc = Round((QFSUC(s1) ^ 2 + QFSUC(s2) ^ 2 - (QFSUC(s1) ^ 2) * (QFSUC(s2) ^ 2)) ^ (1 / 2), 5)
uec = Round(((QFSUE(s1) / tpi) ^ 2 + (QFSUE(s2) / tpi) ^ 2 - ((QFSUE(s1) / tpi) ^ 2) * ((QFSUE(s2) / tpi) ^ 2)) ^ (1 / 2), 5)
vcc = Round(QFSVC(s1) * QFSVC(s2), 5)
vec = Round((QFSVE(s1) / tpi) * (QFSVE(s2) / tpi), 5)
hcc = Round(((1 - QFSUC(s2) ^ 2) * (QFSHC(s1) ^ 2) + (1 - QFSUC(s1) ^ 2) * (QFSHC(s2) ^ 2) - (QFSHC(s1) ^ 2) * (QFSHC(s2) ^ 2)) ^ (1 / 2), 5)
hec = Round(((1 - (QFSUE(s2) / tpi) ^ 2) * ((QFSHE(s1) / tpi) ^ 2) + (1 - (QFSUE(s1) / tpi) ^ 2) * ((QFSHE(s2) / tpi) ^ 2) - ((QFSHE(s1) / tpi) ^ 2) * ((QFSHE(s2) / tpi) ^ 2)) ^ (1 / 2), 5)
QFSADD = "(" & ucc & "ej2p" & uec & "; " & vcc & "ej2p" & vec & "; " & hcc & "ej2p" & hec & ")"
End Function

Function QFSMUL(s1 As String, s2 As String) As String
Dim ucc As Double, uec As Double, tpi As Double, vcc As Double, vec As Double, hcc As Double, hec As Double
Dim i As Integer, pi As Integer
tpi = 2 * WorksheetFunction.pi()
ucc = Round(QFSUC(s1) * QFSUC(s2), 5)
uec = Round((QFSUE(s1) / tpi) * (QFSUE(s2) / tpi), 5)
vcc = Round((QFSVC(s1) ^ 2 + QFSVC(s2) ^ 2 - (QFSVC(s1) ^ 2) * (QFSVC(s2) ^ 2)) ^ (1 / 2), 5)
vec = Round(((QFSVE(s1) / tpi) ^ 2 + (QFSVE(s2) / tpi) ^ 2 - ((QFSVE(s1) / tpi) ^ 2) * ((QFSVE(s2) / tpi) ^ 2)) ^ (1 / 2), 5)
hcc = Round((((1 - QFSVC(s2) ^ 2) * QFSHC(s1) ^ 2) + ((1 - QFSVC(s1) ^ 2) * QFSHC(s2) ^ 2) - (QFSHC(s1) ^ 2) * (QFSHC(s2) ^ 2)) ^ (1 / 2), 5)
hec = Round(((1 - (QFSVE(s2) / tpi) ^ 2) * ((QFSHE(s1) / tpi) ^ 2) + (1 - (QFSVE(s1) / tpi) ^ 2) * ((QFSHE(s2) / tpi) ^ 2) - ((QFSHE(s1) / tpi) ^ 2) * ((QFSHE(s2) / tpi) ^ 2)) ^ (1 / 2), 5)
QFSMUL = "(" & ucc & "ej2p" & uec & "; " & vcc & "ej2p" & vec & "; " & hcc & "ej2p" & hec & ")"
End Function

Function QFSNOM(s As String, c As Double, e As Double) As String
Dim ucc As Double, uec As Double, tpi As Double, vcc As Double, vec As Double, hcc As Double, hec As Double
ucc = Round(QFSUC(s) / c, 5)
uec = Round(QFSUE(s) / e, 5)
vcc = Round(QFSVC(s) / c, 5)
vec = Round(QFSVE(s) / e, 5)
hcc = Round(QFSHC(s) / c, 5)
hec = Round(QFSHE(s) / e, 5)
QFSNOM = "(" & ucc & "ej2p" & uec & "; " & vcc & "ej2p" & vec & "; " & hcc & "ej2p" & hec & ")"
End Function

Function QFSCOE(s As String, k As Double) As String
Dim ucc As Double, uec As Double, tpi As Double, vcc As Double, vec As Double, hcc As Double, hec As Double, a As Double, b As Double, c As Double
Dim i As Integer, pi As Integer
tpi = 2 * WorksheetFunction.pi()
ucc = Round((1 - (1 - QFSUC(s) ^ 2) ^ k) ^ (1 / 2), 5)
uec = Round((1 - ((1 - (QFSUE(s) / tpi) ^ 2) ^ k)) ^ (1 / 2), 5)

vcc = Round(QFSVC(s) ^ k, 5)
vec = Round((QFSVE(s) / tpi) ^ k, 5)

'hcc = Round((((1 - QFSHC(s) ^ 2) ^ k) - ((1 - (QFSUC(s) ^ 2) - (QFSHC(s) ^ 2)) ^ k)) ^ (1 / 2), 3)
a = (1 - QFSHC(s) ^ 2) ^ k
b = (1 - (QFSUC(s) ^ 2) - (QFSHC(s) ^ 2)) ^ k
hcc = Round((a - b) ^ (1 / 2), 5)
'hec = Round((((1 - (QFSHE(s) / tpi) ^ 2) ^ k) - ((1 - ((QFSUE(s) / tpi) ^ 2) - ((QFSHE(s) / tpi) ^ 2)) ^ k)) ^ (1 / 2), 3)
a = (1 - (QFSHE(s) / tpi) ^ 2) ^ k
b = (1 - (QFSUE(s) / tpi) ^ 2 - (QFSHE(s) / tpi) ^ 2) ^ k
hec = Round((a - b) ^ (1 / 2), 5)
QFSCOE = "(" & ucc & "ej2p" & uec & "; " & vcc & "ej2p" & vec & "; " & hcc & "ej2p" & hec & ")"
End Function

Function QFSDEF(s As String) As Double
Dim i As Integer, j As Integer, a As Double, b As Double, tpi As Double
tpi = 2 * WorksheetFunction.pi()
a = QFSUC(s) + QFSHC(s) * (QFSUC(s) / (QFSUC(s) + QFSVC(s)))
b = (QFSUE(s) / tpi) + (QFSVE(s) / tpi) * ((QFSUE(s) / tpi) / ((QFSUE(s) / tpi) + (QFSHE(s) / tpi)))
QFSDEF = Round(a + b, 5)
End Function