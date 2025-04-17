Function NZNTV(s As String) As Double
Dim i As Integer, j As Integer, n As Integer, chia As Integer, vt1 As Integer, vt2 As Integer, demmo As Integer, demngat As Integer, ngat As Integer
Dim a() As Variant
Dim b() As Variant
Dim t As String

n = Len(s)
ReDim a(n)
demmo = 0
demngat = 0

For i = 1 To n
    a(i) = Mid(s, i, 1)
    If a(i) = "(" Then demmo = demmo + 1
    If a(i) = ")" Then demngat = demngat + 1
    If demmo = 1 Then
        vt1 = i
        demmo = 1000
    End If
    If demngat = 1 Then
        vt2 = i
        Exit For
    End If
Next i

For i = vt1 To vt2
    If a(i) = ";" Then
    chia = i
    Exit For
    End If
Next i

NZNTV = CDbl(Mid(s, vt1 + 1, chia - vt1 - 1))
End Function

Function NZNTR(s As String) As Double
Dim i As Integer, j As Integer, n As Integer, chia As Integer, vt1 As Integer, vt2 As Integer, demmo As Integer, demngat As Integer, ngat As Integer
Dim a() As Variant
Dim b() As Variant
Dim t As String

n = Len(s)
ReDim a(n)
demmo = 0
demngat = 0

For i = 1 To n
    a(i) = Mid(s, i, 1)
    If a(i) = "(" Then demmo = demmo + 1
    If a(i) = ")" Then demngat = demngat + 1
    If demmo = 1 Then
        vt1 = i
        demmo = 1000
    End If
    If demngat = 1 Then
        vt2 = i
        Exit For
    End If
Next i

For i = vt1 To vt2
    If a(i) = ";" Then
    chia = i
    Exit For
    End If
Next i

NZNTR = CDbl(Mid(s, chia + 1, vt2 - chia - 1))
End Function

Function NZNIR(s As String) As Double
Dim i As Integer, j As Integer, n As Integer, chia As Integer, vt1 As Integer, vt2 As Integer, demmo As Integer, demngat As Integer, ngat As Integer
Dim a() As Variant
Dim b() As Variant
Dim t As String

n = Len(s)
ReDim a(n)
demmo = 0
demngat = 0

For i = 1 To n
    a(i) = Mid(s, i, 1)
    If a(i) = "(" Then demmo = demmo + 1
    If a(i) = ")" Then demngat = demngat + 1
    If demmo = 2 Then
        vt1 = i
        demmo = 1000
    End If
    If demngat = 2 Then
        vt2 = i
        Exit For
    End If
Next i

For i = vt1 To vt2
    If a(i) = ";" Then
    chia = i
    Exit For
    End If
Next i

NZNIR = CDbl(Mid(s, chia + 1, vt2 - chia - 1))
End Function

Function NZNIV(s As String) As Double
Dim i As Integer, j As Integer, n As Integer, chia As Integer, vt1 As Integer, vt2 As Integer, demmo As Integer, demngat As Integer, ngat As Integer
Dim a() As Variant
Dim b() As Variant
Dim t As String

n = Len(s)
ReDim a(n)
demmo = 0
demngat = 0

For i = 1 To n
    a(i) = Mid(s, i, 1)
    If a(i) = "(" Then demmo = demmo + 1
    If a(i) = ")" Then demngat = demngat + 1
    If demmo = 2 Then
        vt1 = i
        demmo = 1000
    End If
    If demngat = 2 Then
        vt2 = i
        Exit For
    End If
Next i

For i = vt1 To vt2
    If a(i) = ";" Then
    chia = i
    Exit For
    End If
Next i

NZNIV = CDbl(Mid(s, vt1 + 1, chia - vt1 - 1))
End Function

Function NZNFR(s As String) As Double
Dim i As Integer, j As Integer, n As Integer, chia As Integer, vt1 As Integer, vt2 As Integer, demmo As Integer, demngat As Integer, ngat As Integer
Dim a() As Variant
Dim b() As Variant
Dim t As String

n = Len(s)
ReDim a(n)
demmo = 0
demngat = 0

For i = 1 To n
    a(i) = Mid(s, i, 1)
    If a(i) = "(" Then demmo = demmo + 1
    If a(i) = ")" Then demngat = demngat + 1
    If demmo = 3 Then
        vt1 = i
        demmo = 1000
    End If
    If demngat = 3 Then
        vt2 = i
        Exit For
    End If
Next i

For i = vt1 To vt2
    If a(i) = ";" Then
    chia = i
    Exit For
    End If
Next i

NZNFR = CDbl(Mid(s, chia + 1, vt2 - chia - 1))
End Function

Function NZNFV(s As String) As Double
Dim i As Integer, j As Integer, n As Integer, chia As Integer, vt1 As Integer, vt2 As Integer, demmo As Integer, demngat As Integer, ngat As Integer
Dim a() As Variant
Dim b() As Variant
Dim t As String

n = Len(s)
ReDim a(n)
demmo = 0
demngat = 0

For i = 1 To n
    a(i) = Mid(s, i, 1)
    If a(i) = "(" Then demmo = demmo + 1
    If a(i) = ")" Then demngat = demngat + 1
    If demmo = 3 Then
        vt1 = i
        demmo = 1000
    End If
    If demngat = 3 Then
        vt2 = i
        Exit For
    End If
Next i

For i = vt1 To vt2
    If a(i) = ";" Then
    chia = i
    Exit For
    End If
Next i

NZNFV = CDbl(Mid(s, vt1 + 1, chia - vt1 - 1))
End Function

Function NZNADD(s1 As String, s2 As String) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tva As Double, Tra As Double, Iva As Double, Ira As Double, Fra As Double, Fva As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String
Dim g() As Variant
Tv1 = NZNTV(s1)
Tr1 = NZNTR(s1)
Iv1 = NZNIV(s1)
Ir1 = NZNIR(s1)
Fv1 = NZNFV(s1)
Fr1 = NZNFR(s1)

Tv2 = NZNTV(s2)
Tr2 = NZNTR(s2)
Iv2 = NZNIV(s2)
Ir2 = NZNIR(s2)
Fv2 = NZNFV(s2)
Fr2 = NZNFR(s2)

Tva = Round(Tv1 + Tv2 - Tv1 * Tv2, 4)
Tra = Round(Tr1 + Tr2 - Tr1 * Tr2, 4)
Iva = Round(Iv1 * Iv2, 4)
Ira = Round(Ir1 * Ir2, 4)
Fva = Round(Fv1 * Fv2, 4)
Fra = Round(Fr1 * Fr2, 4)

s = ""
s = "[(" & Tva & ";" & Tra & ");" & "(" & Iva & ";" & Ira & ");" & "(" & Fva & ";" & Fra & ")]"
NZNADD = s
End Function

Function NZNMUL(s1 As String, s2 As String) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvm As Double, Trm As Double, Ivm As Double, Irm As Double, Frm As Double, Fvm As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String
Dim g() As Variant
Tv1 = NZNTV(s1)
Tr1 = NZNTR(s1)
Iv1 = NZNIV(s1)
Ir1 = NZNIR(s1)
Fv1 = NZNFV(s1)
Fr1 = NZNFR(s1)

Tv2 = NZNTV(s2)
Tr2 = NZNTR(s2)
Iv2 = NZNIV(s2)
Ir2 = NZNIR(s2)
Fv2 = NZNFV(s2)
Fr2 = NZNFR(s2)

Tvm = Round(Tv1 * Tv2, 4)
Trm = Round(Tr1 * Tr2, 4)
Ivm = Round(Iv1 + Iv2 - Iv1 * Iv2, 4)
Irm = Round(Ir1 + Ir2 - Ir1 * Ir2, 4)
Fvm = Round(Fv1 + Fv2 - Fv1 * Fv2, 4)
Frm = Round(Fr1 + Fr2 - Fr1 * Fr2, 4)

s = ""
s = "[(" & Tvm & ";" & Trm & ");" & "(" & Ivm & ";" & Irm & ");" & "(" & Fvm & ";" & Frm & ")]"
NZNMUL = s
End Function

Function NZNCOE(s1 As String, coe As Double) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvc As Double, Trc As Double, Ivc As Double, Irc As Double, Frc As Double, Fvc As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String
Dim g() As Variant
Tv1 = NZNTV(s1)
Tr1 = NZNTR(s1)
Iv1 = NZNIV(s1)
Ir1 = NZNIR(s1)
Fv1 = NZNFV(s1)
Fr1 = NZNFR(s1)

Tvc = Round(1 - (1 - Tv1) ^ coe, 4)
Trc = Round(1 - (1 - Tr1) ^ coe, 4)
Ivc = Round((Iv1) ^ coe, 4)
Irc = Round((Ir1) ^ coe, 4)
Fvc = Round((Fv1) ^ coe, 4)
Frc = Round((Fr1) ^ coe, 4)

s = ""
s = "[(" & Tvc & ";" & Trc & ");" & "(" & Ivc & ";" & Irc & ");" & "(" & Fvc & ";" & Frc & ")]"
NZNCOE = s
End Function
Function NZNPOW(s1 As String, pow As Double) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvw As Double, Trw As Double, Ivw As Double, Irw As Double, Frw As Double, Fvw As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String
Dim g() As Variant
Tv1 = NZNTV(s1)
Tr1 = NZNTR(s1)
Iv1 = NZNIV(s1)
Ir1 = NZNIR(s1)
Fv1 = NZNFV(s1)
Fr1 = NZNFR(s1)

Tvw = Round((Tv1) ^ pow, 4)
Trw = Round((Tr1) ^ pow, 4)
Ivw = Round(1 - (1 - Iv1) ^ pow, 4)
Irw = Round(1 - (1 - Ir1) ^ pow, 4)
Fvw = Round(1 - (1 - Fv1) ^ pow, 4)
Frw = Round(1 - (1 - Fr1) ^ pow, 4)

s = ""
s = "[(" & Tvw & ";" & Trw & ");" & "(" & Ivw & ";" & Irw & ");" & "(" & Fvw & ";" & Frw & ")]"

NZNPOW = s
End Function

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

Function NZNSCORE(s1 As String) As Double
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvw As Double, Trw As Double, Ivw As Double, Irw As Double, Frw As Double, Fvw As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer, score As Double
Dim s As String
Dim g() As Variant
Tv1 = NZNTV(s1)
Tr1 = NZNTR(s1)
Iv1 = NZNIV(s1)
Ir1 = NZNIR(s1)
Fv1 = NZNFV(s1)
Fr1 = NZNFR(s1)

score = Round((2 + Tv1 * Tr1 - Iv1 * Ir1 - Fv1 * Fr1) / 3, 4)
NZNSCORE = score
End Function

Private Function NZNDISCOS(s1 As String, s2 As String) As Double
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvd As Double, Trd As Double, Ivd As Double, Ird As Double, Frd As Double, Fvd As Double, inside As Double, pi As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String
Dim g() As Variant
Tv1 = NZNTV(s1)
Tr1 = NZNTR(s1)
Iv1 = NZNIV(s1)
Ir1 = NZNIR(s1)
Fv1 = NZNFV(s1)
Fr1 = NZNFR(s1)

Tv2 = NZNTV(s2)
Tr2 = NZNTR(s2)
Iv2 = NZNIV(s2)
Ir2 = NZNIR(s2)
Fv2 = NZNFV(s2)
Fr2 = NZNFR(s2)

inside = 1 / 3 * (Abs(Tv1 - Tv2) + Abs(Iv1 - Iv2) + Abs(Fv1 - Fv2)) + 1 / 3 * (Abs(Tr1 - Tr2) + Abs(Ir1 - Ir2) + Abs(Fr1 - Fr2))

pi = WorksheetFunction.pi()
NZNDISCOS = WorksheetFunction.Cot(pi / 4 + (pi / 8) * inside)
End Function

Function NZNPIS(x As Range) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvpis As Double, Trpis As Double, Ivpis As Double, Irpis As Double, Frpis As Double, Fvpis As Double, inside As Double, pi As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String

n = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)

Tvpis = NZNTV(x(1))
Trpis = NZNTR(x(1))
Ivpis = NZNIV(x(1))
Irpis = NZNIR(x(1))
Fvpis = NZNFV(x(1))
Frpis = NZNFR(x(1))

For i = 2 To n
    Tv1 = NZNTV(x(i))
    Tvpis = Round(WorksheetFunction.Max(Tvpis, Tv1), 4)
    Tr1 = NZNTR(x(i))
    Trpis = Round(WorksheetFunction.Max(Trpis, Tr1), 4)
    
    Iv1 = NZNIV(x(i))
    Ivpis = Round(WorksheetFunction.Min(Ivpis, Iv1), 4)
    Ir1 = NZNIR(x(i))
    Irpis = Round(WorksheetFunction.Min(Irpis, Ir1), 4)
    
    Fv1 = NZNFV(x(i))
    Fvpis = Round(WorksheetFunction.Min(Fvpis, Fv1), 4)
    Fr1 = NZNFR(x(i))
    Frpis = Round(WorksheetFunction.Min(Frpis, Fr1), 4)
Next i

s = ""
s = "[(" & Tvpis & ";" & Trpis & ");" & "(" & Ivpis & ";" & Irpis & ");" & "(" & Fvpis & ";" & Frpis & ")]"

NZNPIS = s
End Function
Function NZNNIS(x As Range) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvpis As Double, Trpis As Double, Ivpis As Double, Irpis As Double, Frpis As Double, Fvpis As Double, inside As Double, pi As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String

n = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)

Tvpis = NZNTV(x(1))
Trpis = NZNTR(x(1))
Ivpis = NZNIV(x(1))
Irpis = NZNIR(x(1))
Fvpis = NZNFV(x(1))
Frpis = NZNFR(x(1))

For i = 2 To n
    Tv1 = NZNTV(x(i))
    Tvpis = Round(WorksheetFunction.Min(Tvpis, Tv1), 4)
    Tr1 = NZNTR(x(i))
    Trpis = Round(WorksheetFunction.Min(Trpis, Tr1), 4)
    
    Iv1 = NZNIV(x(i))
    Ivpis = Round(WorksheetFunction.Max(Ivpis, Iv1), 4)
    Ir1 = NZNIR(x(i))
    Irpis = Round(WorksheetFunction.Max(Irpis, Ir1), 4)
    
    Fv1 = NZNFV(x(i))
    Fvpis = Round(WorksheetFunction.Max(Fvpis, Fv1), 4)
    Fr1 = NZNFR(x(i))
    Frpis = Round(WorksheetFunction.Max(Frpis, Fr1), 4)
Next i

s = ""
s = "[(" & Tvpis & ";" & Trpis & ");" & "(" & Ivpis & ";" & Irpis & ");" & "(" & Fvpis & ";" & Frpis & ")]"

NZNPIS = s
End Function
Private Function NZNDISV(x As Range, PIS As Range, k As Range, p As Double) As Double
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvd As Double, Trd As Double, Ivd As Double, Ird As Double, Frd As Double, Fvd As Double, inside As Double, pi As Double, tv As Double, tr As Double, hs As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer, u As Integer
Dim s1 As String, s2 As String
Dim g() As String

n = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)
m = WorksheetFunction.Max(k.Rows.Count, k.Columns.Count)
tv = 0

For j = 1 To n
    s1 = x(j)
    s2 = PIS(j)

    Tv1 = NZNTV(s1)
    Tr1 = NZNTR(s1)
    Iv1 = NZNIV(s1)
    Ir1 = NZNIR(s1)
    Fv1 = NZNFV(s1)
    Fr1 = NZNFR(s1)
    
    Tv2 = NZNTV(s2)
    Tr2 = NZNTR(s2)
    Iv2 = NZNIV(s2)
    Ir2 = NZNIR(s2)
    Fv2 = NZNFV(s2)
    Fr2 = NZNFR(s2)
    
    tv = tv + k(j) * ((Abs(Tv1 - Tv2) ^ (p)) + ((Abs(Iv1 - Iv2)) ^ (p)) + ((Abs(Fv1 - Fv2)) ^ (p)))
Next j

tv = tv / 3
tv = tv ^ (1 / p)

NZNDISV = tv
End Function

Private Function NZNDISR(x As Range, PIS As Range, k As Range, p As Double) As Double
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvd As Double, Trd As Double, Ivd As Double, Ird As Double, Frd As Double, Fvd As Double, inside As Double, pi As Double, tr As Double, hs As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer, u As Integer
Dim s1 As String, s2 As String
Dim g() As String

n = WorksheetFunction.Max(x.Rows.Count, x.Columns.Count)
m = WorksheetFunction.Max(k.Rows.Count, k.Columns.Count)
tr = 0

For j = 1 To n
    s1 = x(j)
    s2 = PIS(j)

    Tv1 = NZNTV(s1)
    Tr1 = NZNTR(s1)
    Iv1 = NZNIV(s1)
    Ir1 = NZNIR(s1)
    Fv1 = NZNFV(s1)
    Fr1 = NZNFR(s1)
    
    Tv2 = NZNTV(s2)
    Tr2 = NZNTR(s2)
    Iv2 = NZNIV(s2)
    Ir2 = NZNIR(s2)
    Fv2 = NZNFV(s2)
    Fr2 = NZNFR(s2)
    
    tr = tr + k(j) * ((Abs(Tr1 - Tr2) ^ (p)) + ((Abs(Ir1 - Ir2)) ^ (p)) + ((Abs(Fr1 - Fr2)) ^ (p)))
Next j

tr = tr / 3
tr = tr ^ (1 / p)

NZNDISR = tr
End Function

Function NZNDIS(x As Range, PIS As Range, k As Range, p As Double) As Double
NZNDIS = Round((1 / 2) * (NZNDISV(x, PIS, k, p) + NZNDISR(x, PIS, k, p)), 4)
End Function

Function NZNCOM(s1 As String) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim Tv2 As Double, Tr2 As Double, Fv2 As Double, Fr2 As Double, Iv2 As Double, Ir2 As Double
Dim Tvd As Double, Trd As Double, Ivd As Double, Ird As Double, Frd As Double, Fvd As Double, inside As Double, pi As Double, tr As Double, hs As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer, u As Integer
Dim s2 As String

    Tv1 = NZNTV(s1)
    Tr1 = NZNTR(s1)
    Iv1 = NZNIV(s1)
    Iv1 = 1 - Iv1
    Ir1 = NZNIR(s1)
    Ir1 = 1 - Ir1
    Fv1 = NZNFV(s1)
    Fr1 = NZNFR(s1)

s = ""
s = "[(" & Fv1 & ";" & Fr1 & ");" & "(" & Iv1 & ";" & Ir1 & ");" & "(" & Tv1 & ";" & Tr1 & ")]"
NZNCOM = s
End Function

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

Function NZNNOM(s1 As String, p As Double, q As Double) As String
Dim Tv1 As Double, Tr1 As Double, Fv1 As Double, Fr1 As Double, Iv1 As Double, Ir1 As Double
Dim s As String
Dim g() As Variant

Tv1 = NZNTV(s1)
Tv1 = Round(Tv1 / p, 4)
Tr1 = NZNTR(s1)
Tr1 = Round(Tr1 / q, 4)
Iv1 = NZNIV(s1)
Iv1 = Round(Iv1 / p, 4)
Ir1 = NZNIR(s1)
Ir1 = Round(Ir1 / q, 4)
Fv1 = NZNFV(s1)
Fv1 = Round(Fv1 / p, 4)
Fr1 = NZNFR(s1)
Fr1 = Round(Fr1 / q, 4)

s = ""
s = "[(" & Tv1 & ";" & Tr1 & ");" & "(" & Iv1 & ";" & Ir1 & ");" & "(" & Fv1 & ";" & Fr1 & ")]"
NZNNOM = s
End Function

Function MNZNWAA()
MNZNWAA
End Function

Function MNZNWGA()
MNZNWGA
End Function
