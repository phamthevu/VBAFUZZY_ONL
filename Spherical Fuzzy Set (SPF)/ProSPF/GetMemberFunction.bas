Function GET_R(x As String) As Double
Dim mem As Double, non As Double, hes As Double

mem = GET_A(x)
non = GET_B(x)
hes = GET_Y(x)

GET_R = 1 - mem * mem - non * non - hes * hes
End Function

Function SPF_CONVERT_TO_PRO(x As String) As String
Dim mem As Double, non As Double, hes As Double
Dim k1 As Double, k2 As Double, r As Double

mem = GET_A(x)
non = GET_B(x)
hes = GET_Y(x)

k1 = Round(mem / hes, 4)
k2 = Round(non / hes, 4)

r = GET_R(x)

SPF_CONVERT_TO_PRO = "(" & k1 & "; " & k2 & "; " & r & ")"

End Function

Function PRO_CONVERT_TO_SPF(x As String) As String
Dim k1 As Double, k2 As Double, r As Double
Dim mem As Double, non As Double, hes As Double

k1 = GET_A(x)
k2 = GET_B(x)
r = GET_Y(x)

mem = Round(k1 * (((1 - r) / (k1 * k1 + k2 * k2 + 1)) ^ (1 / 2)), 4)
non = Round(k2 * (((1 - r) / (k1 * k1 + k2 * k2 + 1)) ^ (1 / 2)), 4)
hes = Round((((1 - r) / (k1 * k1 + k2 * k2 + 1)) ^ (1 / 2)), 4)

PRO_CONVERT_TO_SPF = "(" & mem & "; " & non & "; " & hes & ")"
End Function