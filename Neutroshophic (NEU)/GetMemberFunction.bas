Function GET_AL(x As String) As Double
Dim pos As Integer, A As String, Num_A As Double
pos = WorksheetFunction.Find(";", x, 1)
A = Mid(x, 2, pos - 2)
Num_A = Round(CDbl(A), 3)
GET_AL = Num_A
End Function

Function GET_BE(x As String) As Double
Dim pos1 As Integer, pos2 As Integer, B As String, Num_B As Double
pos1 = WorksheetFunction.Find(";", x, 1)
pos2 = WorksheetFunction.Find(";", x, pos1 + 1)
B = Mid(x, pos1 + 2, pos2 - pos1 - 2)
Num_B = Round(CDbl(B), 3)
GET_BE = Num_B
End Function

Function GET_YM(x As String) As Double
Dim pos1 As Integer, pos2 As Integer, Y As String, Num_Y As Double, l As Integer
pos1 = WorksheetFunction.Find(";", x, 1)
pos2 = WorksheetFunction.Find(";", x, pos1 + 1)
l = Len(x)
Y = Mid(x, pos2 + 2, l - pos2 - 2)
Num_Y = Round(CDbl(Y), 3)
GET_YM = Num_Y
End Function






