Function GET_POS(x As String) As Double
Dim pos As Integer, P As Double

pos = WorksheetFunction.Find(";", x, 1)
P = Round(CDbl(Mid(x, 2, pos - 2)), 4)

GET_P = P
End Function

Function GET_NAU(x As String) As Double
Dim pos1 As Integer, N As Double, pos2 As Integer

pos1 = WorksheetFunction.Find(";", x, 1)
pos2 = WorksheetFunction.Find(";", x, pos1 + 1)
N = Round(CDbl(Mid(x, pos1 + 2, pos2 - pos1 - 2)), 4)

GET_N = N
End Function

Function GET_NEG(x As String) As Double
Dim pos1 As Integer, NEG As Double, pos2 As Integer, l As Integer

pos1 = WorksheetFunction.Find(";", x, 1)
pos2 = WorksheetFunction.Find(";", x, pos1 + 1)
l = Len(x)
NEG = Round(CDbl(Mid(x, pos2 + 2, l - pos2 - 2)), 4)

GET_NEG = NEG
End Function

Function GET_R(x As String) As Double
Dim R As Double, pos As Double, neu As Double, neg As Double

pos = GET_POS(x)
neu = GET_NEU(x)
neg = GET_NEG(x)

R = 1 - pos - neu - neg

GET_R = Round(R, 4)

End Function