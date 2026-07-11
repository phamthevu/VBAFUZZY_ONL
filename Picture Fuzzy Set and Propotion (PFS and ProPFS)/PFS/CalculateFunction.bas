Function PFS_SUM(x1 As String, x2 As String) As String
Dim pos1 As Double, pos2 As Double, neu1 As Double, neu2 As Double, neg1 As Double, neg2 As Double
Dim pos As Double, neu As Double, neg As Double

pos1 = GET_POS(x1)
pos2 = GET_POS(x2)
neu1 = GET_NEU(x1)
neu2 = GET_NEU(x2)
neg1 = GET_NEG(x1)
neg2 = GET_NEG(x2)

pos = Round(pos1 + pos2 - pos1 * pos2, 4)
neu = Round(neu1 * neu2, 4)
neg = Round(neg1 * neg2, 4)

PFS_SUM = "(" & pos & "; " & neu & "; " & neg & ")"

End Function

Function PFS_MUL(x1 As String, x2 As String) As String
Dim pos1 As Double, pos2 As Double, neu1 As Double, neu2 As Double, neg1 As Double, neg2 As Double
Dim pos As Double, neu As Double, neg As Double

pos1 = GET_POS(x1)
pos2 = GET_POS(x2)
neu1 = GET_NEU(x1)
neu2 = GET_NEU(x2)
neg1 = GET_NEG(x1)
neg2 = GET_NEG(x2)

pos = Round(pos1 * pos2, 4)
neu = Round(neu1 + neu2 - neu1 * neu2, 4)
neg = Round(neg1 + neg2 - neg1 * neg2, 4)

PFS_MUL = "(" & pos & "; " & neu & "; " & neg & ")"

End Function

Function PFS_COE(x1 As String, n As Double) As String
Dim pos1 As Double, pos2 As Double, neu1 As Double, neu2 As Double, neg1 As Double, neg2 As Double
Dim pos As Double, neu As Double, neg As Double

pos1 = GET_POS(x1)
neu1 = GET_NEU(x1)
neg1 = GET_NEG(x1)


pos = Round(1 - ((1 - pos1) ^ n), 4)
neu = Round(neu1 ^ n, 4)
neg = Round(neg1 ^ n, 4)

PFS_COE = "(" & pos & "; " & neu & "; " & neg & ")"
End Function

Function PFS_POW(x1 As String, n As Double) As String
Dim pos1 As Double, pos2 As Double, neu1 As Double, neu2 As Double, neg1 As Double, neg2 As Double
Dim pos As Double, neu As Double, neg As Double

pos1 = GET_POS(x1)
neu1 = GET_NEU(x1)
neg1 = GET_NEG(x1)


pos = Round(pos1 ^ n, 4)
neu = Round(1 - ((1 - neu1) ^ n), 4)
neg = Round(1 - ((1 - neg1) ^ n), 4)

PFS_POW = "(" & pos & "; " & neu & "; " & neg & ")"
End Function