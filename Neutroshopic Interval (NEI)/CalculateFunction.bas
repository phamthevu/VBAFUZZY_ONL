Function NEIADD(s1 As String, s2 As String) As String
Dim Tl1 As Double, Tu1 As Double, Fl1 As Double, Fu1 As Double, Il1 As Double, Iu1 As Double
Dim Tl2 As Double, Tu2 As Double, Fl2 As Double, Fu2 As Double, Il2 As Double, Iu2 As Double
Dim Tla As Double, Tua As Double, Fla As Double, Fua As Double, Ila As Double, Iua As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer

Dim s As String

Tl1 = NEITL(s1)
Tu1 = NEITU(s1)
Il1 = NEIIL(s1)
Iu1 = NEIIU(s1)
Fl1 = NEIFL(s1)
Fu1 = NEIFU(s1)

Tl2 = NEITL(s2)
Tu2 = NEITU(s2)
Il2 = NEIIL(s2)
Iu2 = NEIIU(s2)
Fl2 = NEIFL(s2)
Fu2 = NEIFU(s2)

Tla = Round(Tl1 + Tl2 - Tl1 * Tl2, 4)
Tua = Round(Tu1 + Tu2 - Tu1 * Tu2, 4)
Ila = Round(Il1 * Il2, 4)
Iua = Round(Iu1 * Iu2, 4)
Fla = Round(Fl1 * Fl2, 4)
Fua = Round(Fu1 * Fu2, 4)

s = ""
s = "[(" & Tla & ";" & Tua & ");" & "(" & Ila & ";" & Iua & ");" & "(" & Fla & ";" & Fua & ")]"
NEIADD = s
End Function

Function NEIMUL(s1 As String, s2 As String) As String
Dim Tl1 As Double, Tu1 As Double, Fl1 As Double, Fu1 As Double, Il1 As Double, Iu1 As Double
Dim Tl2 As Double, Tu2 As Double, Fl2 As Double, Fu2 As Double, Il2 As Double, Iu2 As Double
Dim Tla As Double, Tua As Double, Fla As Double, Fua As Double, Ila As Double, Iua As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer

Dim s As String

Tl1 = NEITL(s1)
Tu1 = NEITU(s1)
Il1 = NEIIL(s1)
Iu1 = NEIIU(s1)
Fl1 = NEIFL(s1)
Fu1 = NEIFU(s1)

Tl2 = NEITL(s2)
Tu2 = NEITU(s2)
Il2 = NEIIL(s2)
Iu2 = NEIIU(s2)
Fl2 = NEIFL(s2)
Fu2 = NEIFU(s2)

Tla = Round(Tl1 * Tl2, 4)
Tua = Round(Tu1 * Tu2, 4)
Ila = Round(Il1 + Il2 - Il1 * Il2, 4)
Iua = Round(Iu1 + Iu2 - Iu1 * Iu2, 4)
Fla = Round(Fl1 + Fl2 - Fl1 * Fl2, 4)
Fua = Round(Fu1 + Fu2 - Fu1 * Fu2, 4)

s = ""
s = "[(" & Tla & ";" & Tua & ");" & "(" & Ila & ";" & Iua & ");" & "(" & Fla & ";" & Fua & ")]"
NEIMUL = s
End Function

Function NEICOE(s1 As String, coe As Double) As String
Dim Tl1 As Double, Tu1 As Double, Fl1 As Double, Fu1 As Double, Il1 As Double, Iu1 As Double
Dim Tlc As Double, Tuc As Double, Flc As Double, Fuc As Double, Ilc As Double, Iuc As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer

Dim s As String

Tl1 = NEITL(s1)
Tu1 = NEITU(s1)
Il1 = NEIIL(s1)
Iu1 = NEIIU(s1)
Fl1 = NEIFL(s1)
Fu1 = NEIFU(s1)

Tvc = Round(1 - (1 - Tl1) ^ coe, 4)
Trc = Round(1 - (1 - Tu1) ^ coe, 4)
Ivc = Round((Il1) ^ coe, 4)
Irc = Round((Iu1) ^ coe, 4)
Fvc = Round((Fl1) ^ coe, 4)
Frc = Round((Fu1) ^ coe, 4)

s = ""
s = "[(" & Tvc & ";" & Trc & ");" & "(" & Ivc & ";" & Irc & ");" & "(" & Fvc & ";" & Frc & ")]"
NEICOE = s
End Function

Function NEIPOW(s1 As String, coe As Double) As String
Dim Tl1 As Double, Tu1 As Double, Fl1 As Double, Fu1 As Double, Il1 As Double, Iu1 As Double
Dim Tlc As Double, Tuc As Double, Flc As Double, Fuc As Double, Ilc As Double, Iuc As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer

Dim s As String

Tl1 = NEITL(s1)
Tu1 = NEITU(s1)
Il1 = NEIIL(s1)
Iu1 = NEIIU(s1)
Fl1 = NEIFL(s1)
Fu1 = NEIFU(s1)

Tvc = Round((Tl1) ^ coe, 4)
Trc = Round((Tu1) ^ coe, 4)
Ivc = Round(1 - (1 - Il1) ^ coe, 4)
Irc = Round(1 - (1 - Iu1) ^ coe, 4)
Fvc = Round(1 - (1 - Fl1) ^ coe, 4)
Frc = Round(1 - (1 - Fu1) ^ coe, 4)

s = ""
s = "[(" & Tvc & ";" & Trc & ");" & "(" & Ivc & ";" & Irc & ");" & "(" & Fvc & ";" & Frc & ")]"
NEIPOW = s
End Function

Function NEISCORE(s1 As String) As Double
Dim Tl1 As Double, Tu1 As Double, Fl1 As Double, Fu1 As Double, Il1 As Double, Iu1 As Double

Tl1 = NEITL(s1)
Tu1 = NEITU(s1)
Il1 = NEIIL(s1)
Iu1 = NEIIU(s1)
Fl1 = NEIFL(s1)
Fu1 = NEIFU(s1)

NEISCORE = 1 / 4 * (2 + Tl1 + Tu1 - 2 * (Il1 + Iu1) - 2 * (Fl1 + Fu1))
End Function

Function NEIPIS(x As Range) As String
Dim Tu1 As Double, Tl1 As Double, Fu1 As Double, Fl1 As Double, Iu1 As Double, Il1 As Double
Dim TUpis As Double, TLpis As Double, IUpis As Double, ILpis As Double, FUpis As Double, FLpis As Double, inside As Double, pi As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String

n = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)

TUpis = NEITU(x(1))
TLpis = NEITL(x(1))
IUpis = NEIIU(x(1))
ILpis = NEIIL(x(1))
FUpis = NEIFU(x(1))
FLpis = NEIFL(x(1))

For i = 2 To n
    Tu1 = NEITU(x(i))
    TUpis = Round(WorksheetFunction.Max(TUpis, Tu1), 4)
    Tl1 = NEITL(x(i))
    TLpis = Round(WorksheetFunction.Max(TLpis, Tl1), 4)
    
    Iu1 = NEIIU(x(i))
    IUpis = Round(WorksheetFunction.Min(IUpis, Iu1), 4)
    Il1 = NEIIL(x(i))
    ILpis = Round(WorksheetFunction.Min(ILpis, Il1), 4)
    
    Fu1 = NEIFU(x(i))
    FUpis = Round(WorksheetFunction.Min(FUpis, Fu1), 4)
    Fl1 = NEIFL(x(i))
    FLpis = Round(WorksheetFunction.Min(FLpis, Fl1), 4)
Next i

s = ""
s = "[(" & TLpis & ";" & TUpis & ");" & "(" & ILpis & ";" & IUpis & ");" & "(" & FLpis & ";" & FUpis & ")]"

NEIPIS = s
End Function


Function NEINIS(x As Range) As String
Dim Tu1 As Double, Tl1 As Double, Fu1 As Double, Fl1 As Double, Iu1 As Double, Il1 As Double
Dim TUpis As Double, TLpis As Double, IUpis As Double, ILpis As Double, FUpis As Double, FLpis As Double, inside As Double, pi As Double
Dim i As Integer, j As Integer, m As Integer, n As Integer
Dim s As String

n = WorksheetFunction.Max(x.Columns.Count, x.Rows.Count)

TUpis = NEITU(x(1))
TLpis = NEITL(x(1))
IUpis = NEIIU(x(1))
ILpis = NEIIL(x(1))
FUpis = NEIFU(x(1))
FLpis = NEIFL(x(1))

For i = 2 To n
    Tu1 = NEITU(x(i))
    TUpis = Round(WorksheetFunction.Min(TUpis, Tu1), 4)
    Tl1 = NEITL(x(i))
    TLpis = Round(WorksheetFunction.Min(TLpis, Tl1), 4)
    
    Iu1 = NEIIU(x(i))
    IUpis = Round(WorksheetFunction.Max(IUpis, Iu1), 4)
    Il1 = NEIIL(x(i))
    ILpis = Round(WorksheetFunction.Max(ILpis, Il1), 4)
    
    Fu1 = NEIFU(x(i))
    FUpis = Round(WorksheetFunction.Max(FUpis, Fu1), 4)
    Fl1 = NEIFL(x(i))
    FLpis = Round(WorksheetFunction.Max(FLpis, Fl1), 4)
Next i

s = ""
s = "[(" & TLpis & ";" & TUpis & ");" & "(" & ILpis & ";" & IUpis & ");" & "(" & FLpis & ";" & FUpis & ")]"

NEINIS = s
End Function

Function NEIDIS(s1 As Range, s2 As Range, w As Range, pow As Double) As Double
Dim i As Integer, j As Integer, k As Integer, n As Integer, temp As Double, hso As Double
Dim Tu1 As Double, Tl1 As Double, Fu1 As Double, Fl1 As Double, Iu1 As Double, Il1 As Double
Dim Tu2 As Double, Tl2 As Double, Fu2 As Double, Fl2 As Double, Iu2 As Double, Il2 As Double

n = WorksheetFunction.Max(s1.Columns.Count, s1.Rows.Count)
hso = 1 / (6 * n)
temp = 0

For i = 1 To n
    Tl1 = NEITL(s1(i))
    Tu1 = NEITU(s1(i))
    Il1 = NEIIL(s1(i))
    Iu1 = NEIIU(s1(i))
    Fl1 = NEIFL(s1(i))
    Fu1 = NEIFU(s1(i))
    
    Tl2 = NEITL(s2(i))
    Tu2 = NEITU(s2(i))
    Il2 = NEIIL(s2(i))
    Iu2 = NEIIU(s2(i))
    Fl2 = NEIFL(s2(i))
    Fu2 = NEIFU(s2(i))
    
    temp = temp + (Abs(Tl1 - Tl2) ^ pow + Abs(Tu1 - Tu2) ^ pow + Abs(Il1 - Il2) ^ pow + Abs(Iu1 - Iu2) ^ pow + Abs(Fl1 - Fl2) ^ pow + Abs(Fu1 - Fu2) ^ pow) * w(i)
Next i

NEIDIS = Round((temp * hso) ^ (1 / pow), 4)

End Function