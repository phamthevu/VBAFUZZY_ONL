Function NZNWAA(x As Range) As String
Dim i As Integer, j As Integer, k As Integer, ub As Integer, lb As Integer
Dim t As String
t = ""
k = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)
t = x(1)
For i = 2 To k
    t = NZNADD(t, x(i))
Next i
NZNWAA = t
End Function

Function NZNWGA(x As Range) As String
Dim i As Integer, j As Integer, k As Integer, ub As Integer, lb As Integer
Dim t As String
t = ""
k = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)
t = x(1)
For i = 2 To k
    t = NZNMUL(t, x(i))
Next i
NZNWGA = t
End Function