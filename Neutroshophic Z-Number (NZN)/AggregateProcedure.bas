Sub SuMNZNWGA()
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvd As Double, Trd As Double, Ivd As Double, Ird As Double, Frd As Double, Fvd As Double, inside As Double, pi As Double, tr As Double, hs As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer, u As Integer, dem As Integer
Dim s2 As String, s1 As String
Dim s As String, k As Integer, x As Integer, y As Integer, e As Integer, r As Integer, c As Integer
s = InputBox("Series name")
k = InputBox("Num experts")
n = InputBox("Num rows")
m = InputBox("Num cols")
r = InputBox("First row of data")
c = InputBox("First col of data")
MsgBox ("Please select start cells")

x = Selection.Row
y = Selection.Column

For i = r To n + r - 1
    For j = c To c + m - 1
        sh = s & 1
        s1 = Sheets(sh).Cells(i, j)
            For e = 2 To k
                sh = s & e
                s2 = Sheets(sh).Cells(i, j)
                s1 = NZNMUL(s1, s2)
            Next e
        Cells(i - r + x, j - c + y) = s1
    Next j
Next i
End Sub

Sub SuMNZNWAA()
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvd As Double, Trd As Double, Ivd As Double, Ird As Double, Frd As Double, Fvd As Double, inside As Double, pi As Double, tr As Double, hs As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer, u As Integer, dem As Integer
Dim s2 As String, s1 As String
Dim s As String, k As Integer, x As Integer, y As Integer, e As Integer, r As Integer, c As Integer
s = InputBox("Series name")
k = InputBox("Num experts")
n = InputBox("Num rows")
m = InputBox("Num cols")
r = InputBox("First row of data")
c = InputBox("First col of data")
MsgBox ("Please select start cells")

x = Selection.Row
y = Selection.Column

For i = r To n + r - 1
    For j = c To c + m - 1
        sh = s & 1
        s1 = Sheets(sh).Cells(i, j)
            For e = 2 To k
                sh = s & e
                s2 = Sheets(sh).Cells(i, j)
                s1 = NZNADD(s1, s2)
            Next e
        Cells(i - r + x, j - c + y) = s1
    Next j
Next i
End Sub