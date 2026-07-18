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

Function PRO_SPF_SUM(x1 As String, x2 As String) As String
Dim pfs_x1 As String, pfs_x2 As String, sum As String

pfs_x1 = PRO_CONVERT_TO_SPF(x1)
pfs_x2 = PRO_CONVERT_TO_SPF(x2)

sum = SPFSUM(pfs_x1, pfs_x2)

PRO_SPF_SUM = sum
End Function

Function PRO_SPF_MUL(x1 As String, x2 As String) As String
Dim pfs_x1 As String, pfs_x2 As String, mul As String

pfs_x1 = PRO_CONVERT_TO_SPF(x1)
pfs_x2 = PRO_CONVERT_TO_SPF(x2)

mul = SPFMUL(pfs_x1, pfs_x2)

PRO_SPF_MUL = mul
End Function

Function PRO_SPF_COE(x1 As String, n As Double) As String
Dim pfs_x1 As String, pfs_x2 As String, coe As String

pfs_x1 = PRO_CONVERT_TO_SPF(x1)

coe = SPFCOE(pfs_x1, n)

PRO_SPF_COE = coe
End Function

Function PRO_SPF_POW(x1 As String, n As Double) As String
Dim pfs_x1 As String, pfs_x2 As String, pow As String

pfs_x1 = PRO_CONVERT_TO_SPF(x1)

pow = SPFPOW(pfs_x1, n)

PRO_SPF_POW = pow
End Function