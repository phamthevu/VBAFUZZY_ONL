Function PFS_CONVERT_TO_PRO(x As String) As String
Dim pos As Double, neu As Double, neg As Double
Dim k1 As Double, k2 As Double, r As Double

pos = GET_POS(x)
neu = GET_NEU(x)
neg = GET_NEG(x)

k1 = Round(pos / neu, 4)
k2 = Round(neg / neu, 4)

r = GET_R(x)

PFS_CONVERT_TO_PRO = "(" & k1 & "; " & k2 & "; " & r & ")"


End Function

Function PRO_CONVERT_TO_PFS(x As String) As String
Dim k1 As Double, k2 As Double, r As Double
Dim pos As Double, neu As Double, neg As Double

k1 = GET_POS(x)
k2 = GET_NEU(x)
r = GET_NEG(x)

neu = Round((1 - r) / (1 + k1 + k2), 4)
pos = Round(neu * k1, 4)
neg = Round(neu * k2, 4)

PRO_CONVERT_TO_PFS = "(" & pos & "; " & neu & "; " & neg & ")"
End Function

Function PRO_TRUNC_R(x As String) As String
Dim k1 As Double, k2 As Double

k1 = GET_POS(x)
k2 = GET_NEU(x)

PRO_TRUNC_R = "(" & k1 & "; " & k2 & ")"
End Function