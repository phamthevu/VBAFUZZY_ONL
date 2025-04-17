Function SPFSUM(s1 As String, s2 As String) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A2 As Double, B2 As Double, Y2 As Double, A  As Double, B As Double, C As Double, res As String

A1 = GET_A(s1)
B1 = GET_B(s1)
Y1 = GET_Y(s1)
A2 = GET_A(s2)
B2 = GET_B(s2)
Y2 = GET_Y(s2)

A = Round(((A1 ^ 2) + (A2 ^ 2) - (A1 ^ 2) * (A2 ^ 2)) ^ (1 / 2), 3)
B = Round(B1 * B2, 3)
y = Round(((1 - A2 ^ 2) * (Y1 ^ 2) + (1 - A1 ^ 2) * (Y2 ^ 2) - (Y1 ^ 2) * (Y2 ^ 2)) ^ (1 / 2), 3)

SPFSUM = "(" & A & "; " & B & "; " & y & ")"
End Function

Function SPFMUL(s1 As String, s2 As String) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A2 As Double, B2 As Double, Y2 As Double, A  As Double, B As Double, C As Double, res As String

A1 = GET_A(s1)
B1 = GET_B(s1)
Y1 = GET_Y(s1)
A2 = GET_A(s2)
B2 = GET_B(s2)
Y2 = GET_Y(s2)

A = Round(A1 * A2, 3)
B = Round(((B1 ^ 2) + (B2 ^ 2) - (B1 ^ 2) * (B2 ^ 2)) ^ (1 / 2), 3)
y = Round(((1 - B2 ^ 2) * (Y1 ^ 2) + (1 - B1 ^ 2) * (Y2 ^ 2) - (Y1 ^ 2) * (Y2 ^ 2)) ^ (1 / 2), 3)

SPFMUL = "(" & A & "; " & B & "; " & y & ")"
End Function

Function SPFCOE(s1 As String, n As Double) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A2 As Double, B2 As Double, Y2 As Double, A  As Double, B As Double, C As Double, res As String

A1 = GET_A(s1)
B1 = GET_B(s1)
Y1 = GET_Y(s1)


A = Round((1 - (1 - A1 ^ 2) ^ n) ^ (1 / 2), 3)
B = Round(B1 ^ n, 3)
y = Round((((1 - (A1 ^ 2)) ^ n) - ((1 - (A1 ^ 2) - (Y1 ^ 2)) ^ n)) ^ (1 / 2), 3)

SPFCOE = "(" & A & "; " & B & "; " & y & ")"
End Function

Function SPFPOW(s1 As String, n As Double) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A2 As Double, B2 As Double, Y2 As Double, A  As Double, B As Double, C As Double, res As String

A1 = GET_A(s1)
B1 = GET_B(s1)
Y1 = GET_Y(s1)


A = Round(A1 ^ n, 3)
B = Round((1 - (1 - B1 ^ 2) ^ n) ^ (1 / 2), 3)
y = Round((((1 - (B1 ^ 2)) ^ n) - ((1 - (B1 ^ 2) - (Y1 ^ 2)) ^ n)) ^ (1 / 2), 3)

SPFPOW = "(" & A & "; " & B & "; " & y & ")"
End Function
Function SPFDEF(s1 As String) As Double
Dim A1 As Double, B1 As Double, Y1 As Double

A1 = GET_A(s1)
B1 = GET_B(s1)
Y1 = GET_Y(s1)

SPFDEF = ((2 * A1 - Y1) ^ 2) - ((B1 - Y1) ^ 2)
End Function