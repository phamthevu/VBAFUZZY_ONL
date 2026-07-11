Function PFS_PWFA(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = PFS_COE(s(1), w(1))
For i = 2 To k
    turn = PFS_COE(s(i), w(i))
    res = PFS_SUM(res, turn)
Next i

PFS_PWFA = res
End Function

Function PFS_PWFG(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = PFS_POW(s(1), w(1))
For i = 2 To k
    turn = PFS_POW(s(i), w(i))
    res = PFS_MUL(res, turn)
Next i

PFS_PWFG = res
End Function