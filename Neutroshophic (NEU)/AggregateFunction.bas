Function NEUSWAM(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = NEUCOE(s(1), w(1))
For i = 2 To k
    turn = NEUCOE(s(i), w(i))
    res = NEUADD(res, turn)
Next i

NEUSWAM = res
End Function

Function NEUSWGM(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = NEUPOW(s(1), w(1))
For i = 2 To k
    turn = NEUPOW(s(i), w(i))
    res = NEUMUL(res, turn)
Next i

NEUSWGM = res
End Function