Function NEIWAA(x As Range) As String
Dim i As Integer, j As Integer, k As Integer, ub As Integer, lb As Integer
Dim t As String
t = ""
k = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)
t = x(1)
For i = 2 To k
    t = NEIADD(t, x(i))
Next i
NEIWAA = t
End Function

Function NEIWGA(x As Range) As String
Dim i As Integer, j As Integer, k As Integer, ub As Integer, lb As Integer
Dim t As String
t = ""
k = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)
t = x(1)
For i = 2 To k
    t = NEIMUL(t, x(i))
Next i
NEIWGA = t
End Function

Function NEIWAAW(x As Range, w As Range) As String
Dim i As Integer, j As Integer, kx As Integer, kw As Integer
Dim t As String, temp As String

t = ""
kx = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)
kw = WorksheetFunction.Max(w.Columns.Count, w.Rows.Count)

t = NEICOE(x(1), w(1))
For i = 2 To kx
    temp = NEICOE(x(i), w(i))
    t = NEIADD(t, temp)
Next i
NEIWAAW = t
End Function

Function NEIWGAW(x As Range, w As Range) As String
Dim i As Integer, j As Integer, kx As Integer, kw As Integer
Dim t As String, temp As String

t = ""
kx = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)
kw = WorksheetFunction.Max(w.Columns.Count, w.Rows.Count)

t = NEIPOW(x(1), w(1))
For i = 2 To kx
    temp = NEIPOW(x(i), w(i))
    t = NEIMUL(t, temp)
Next i
NEIWGAW = t
End Function