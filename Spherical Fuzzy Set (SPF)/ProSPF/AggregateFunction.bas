Function PRO_SPF_SWAM(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = PRO_SPF_COE(s(1), w(1))
For i = 2 To k
    turn = PRO_SPF_COE(s(i), w(i))
    res = SPFSUM(res, turn)
Next i

PRO_SPF_SWAM = res
End Function

Function PRO_SPF_SWAG(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = PRO_SPF_POW(s(1), w(1))
For i = 2 To k
    turn = PRO_SPF_POW(s(i), w(i))
    res = SPFMUL(res, turn)
Next i

PRO_SPF_SWAG = res
End Function