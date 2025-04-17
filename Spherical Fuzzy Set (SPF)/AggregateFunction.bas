Function SPFSWAM(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = SPFCOE(s(1), w(1))
For i = 2 To k
    turn = SPFCOE(s(i), w(i))
    res = SPFSUM(res, turn)
Next i

SPFSWAM = res
End Function

Function SPFSWGM(s As Range, w As Range) As String
Dim k As Integer, i As Integer, j As Integer, res As String, turn As String

k = WorksheetFunction.Max(s.Rows.Count, s.Columns.Count)
res = SPFPOW(s(1), w(1))
For i = 2 To k
    turn = SPFPOW(s(i), w(i))
    res = SPFMUL(res, turn)
Next i

SPFSWGM = res
End Function