Function WGAOBC1(x As Range) As String
Dim t As String, j As Integer, n As Integer
n = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)
t = x(1)
For i = 2 To n
    t = MULBC(t, x(i))
Next i
WGAOBC1 = t
End Function

Function WGAOBC2(x As Range, y As Range) As String
Dim t As String, j As Integer, n As Integer, k As String
n = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)
t = POWBC(x(1), y(1))
For i = 2 To n
    k = POWBC(x(i), y(i))
    t = MULBC(t, k)
Next i
WGAOBC2 = t
End Function