Function GET_A(x As String) As Double
Dim pos As Integer, A As String, Num_A As Double
pos = WorksheetFunction.Find(";", x, 1)
A = Mid(x, 2, pos - 2)
Num_A = Round(CDbl(A), 3)
GET_A = Num_A
End Function

Function GET_B(x As String) As Double
Dim pos1 As Integer, pos2 As Integer, B As String, Num_B As Double
pos1 = WorksheetFunction.Find(";", x, 1)
pos2 = WorksheetFunction.Find(";", x, pos1 + 1)
B = Mid(x, pos1 + 2, pos2 - pos1 - 2)
Num_B = Round(CDbl(B), 3)
GET_B = Num_B
End Function

Function GET_Y(x As String) As Double
Dim pos1 As Integer, pos2 As Integer, y As String, Num_Y As Double, l As Integer
pos1 = WorksheetFunction.Find(";", x, 1)
pos2 = WorksheetFunction.Find(";", x, pos1 + 1)
l = Len(x)
y = Mid(x, pos2 + 2, l - pos2 - 2)
Num_Y = Round(CDbl(y), 3)
GET_Y = Num_Y
End Function








