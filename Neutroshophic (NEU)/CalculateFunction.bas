Function NEUADD(s1 As String, s2 As String) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A2 As Double, B2 As Double, Y2 As Double, A As Double, B As Double, Y As Double

A1 = GET_AL(s1)
B1 = GET_BE(s1)
Y1 = GET_YM(s1)
A2 = GET_AL(s2)
B2 = GET_BE(s2)
Y2 = GET_YM(s2)

A = Round(A1 + A2 - A1 * A2, 3)
B = Round(B1 * B2, 3)
Y = Round(Y1 * Y2, 3)

NEUADD = "(" & A & "; " & B & "; " & Y & ")"

End Function

Function NEUMUL(s1 As String, s2 As String) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A2 As Double, B2 As Double, Y2 As Double, A As Double, B As Double, Y As Double

A1 = GET_AL(s1)
B1 = GET_BE(s1)
Y1 = GET_YM(s1)
A2 = GET_AL(s2)
B2 = GET_BE(s2)
Y2 = GET_YM(s2)

A = Round(A1 * A2, 3)
B = Round(B1 + B2 - B1 * B2, 3)
Y = Round(Y1 + Y2 - Y1 * Y2, 3)

NEUMUL = "(" & A & "; " & B & "; " & Y & ")"

End Function

Function NEUCOE(s1 As String, x As Double) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A As Double, B As Double, Y As Double

A1 = GET_AL(s1)
B1 = GET_BE(s1)
Y1 = GET_YM(s1)

A = Round(1 - ((1 - A1) ^ x), 3)
B = Round((B1 ^ x), 3)
Y = Round((Y1 ^ x), 3)

NEUCOE = "(" & A & "; " & B & "; " & Y & ")"

End Function

Function NEUPOW(s1 As String, x As Double) As String
Dim A1 As Double, B1 As Double, Y1 As Double, A As Double, B As Double, Y As Double

A1 = GET_AL(s1)
B1 = GET_BE(s1)
Y1 = GET_YM(s1)

A = Round((A1 ^ x), 3)
B = Round(1 - ((1 - B1) ^ x), 3)
Y = Round(1 - ((1 - Y1) ^ x), 3)

NEUPOW = "(" & A & "; " & B & "; " & Y & ")"

End Function