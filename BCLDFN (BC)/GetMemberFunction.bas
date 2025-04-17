Function PGBC(x As String) As String
Dim i As Integer, j As Integer
i = WorksheetFunction.Find("[", x)
j = WorksheetFunction.Find("]", x)
PGBC = Mid(x, i, j - i + 1)
End Function

Function NGBC(x As String) As String
Dim i As Integer, j As Integer
i1 = WorksheetFunction.Find("[", x)
j1 = WorksheetFunction.Find("]", x)
i = WorksheetFunction.Find("[", x, i1 + 1)
j = WorksheetFunction.Find("]", x, j1 + 1)
NGBC = Mid(x, i, j - i + 1)
End Function

Function CSAHIGHBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dau = 2
dem = 0
i = 0
While dem < 1
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
cuoi = i
s = Trim(Mid(x, dau, cuoi - dau))

j = WorksheetFunction.Find("e", s)
t = Mid(s, 1, j - 1)
CSAHIGHBC = CDbl(t)
End Function

Function PSAHIGHBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dau = 2
dem = 0
i = 0
While dem < 1
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
cuoi = i
s = Trim(Mid(x, dau, cuoi - dau))

j = WorksheetFunction.Find("e", s)
t = Mid(s, j + 6, Len(s) - j - 6)
PSAHIGHBC = CDbl(t)
End Function

Function CSALOWBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dem = 0
i = 0
While dem < 1
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
dau = i + 2
cuoi = WorksheetFunction.Find(";", x, dau + 1)
s = Trim(Mid(x, dau, cuoi - dau))
j = WorksheetFunction.Find("e", s)
t = Mid(s, 1, j - 1)
CSALOWBC = CDbl(t)
End Function

Function PSALOWBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dem = 0
i = 0
While dem < 1
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
dau = i + 2
cuoi = WorksheetFunction.Find(";", x, dau + 1)
s = Trim(Mid(x, dau, cuoi - dau))
j = WorksheetFunction.Find("e", s)
t = Mid(s, j + 6, Len(s) - j - 6)
PSALOWBC = CDbl(t)
End Function

Function CDSAHIGHBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dem = 0
i = 0
While dem < 2
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
dau = i + 2
cuoi = WorksheetFunction.Find(";", x, dau + 1)
s = Trim(Mid(x, dau, cuoi - dau))
j = WorksheetFunction.Find("e", s)
t = Mid(s, 1, j - 1)
CDSAHIGHBC = CDbl(t)
End Function

Function PDSAHIGHBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dem = 0
i = 0
While dem < 2
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
dau = i + 2
cuoi = WorksheetFunction.Find(";", x, dau + 1)
s = Trim(Mid(x, dau, cuoi - dau))
j = WorksheetFunction.Find("e", s)
t = Mid(s, j + 6, Len(s) - j - 6)
PDSAHIGHBC = CDbl(t)
End Function

Function CDSALOWBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dem = 0
i = 0
While dem < 3
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
dau = i + 2
cuoi = Len(x)
s = Trim(Mid(x, dau, cuoi - dau))
j = WorksheetFunction.Find("e", s)
t = Mid(s, 1, j - 1)
CDSALOWBC = CDbl(t)
End Function

Function PDSALOWBC(x As String) As Double
Dim i As Integer, j As Integer, dau As Integer, cuoi As Integer, s As String, t As String
dem = 0
i = 0
While dem < 3
    i = i + 1
    If Mid(x, i, 1) = ";" Then
        dem = dem + 1
    End If
Wend
dau = i + 2
cuoi = Len(x)
s = Trim(Mid(x, dau, cuoi - dau))
j = WorksheetFunction.Find("e", s)
t = Mid(s, j + 6, Len(s) - j - 6)
PDSALOWBC = CDbl(t)
End Function





Function SCOREBC(x As String) As Double
Dim i As Integer
'Positive grade
pgchs1 = CSAHIGHBC(PGBC(x))
pgphs1 = PSAHIGHBC(PGBC(x))
pgcls1 = CSALOWBC(PGBC(x))
pgpls1 = PSALOWBC(PGBC(x))
pgchds1 = CDSAHIGHBC(PGBC(x))
pgphds1 = PDSAHIGHBC(PGBC(x))
pgclds1 = CDSALOWBC(PGBC(x))
pgplds1 = PDSALOWBC(PGBC(x))
'Negative grade
ngchs1 = CSAHIGHBC(NGBC(x))
ngphs1 = PSAHIGHBC(NGBC(x))
ngcls1 = CSALOWBC(NGBC(x))
ngpls1 = PSALOWBC(NGBC(x))
ngchds1 = CDSAHIGHBC(NGBC(x))
ngphds1 = PDSAHIGHBC(NGBC(x))
ngclds1 = CDSALOWBC(NGBC(x))
ngplds1 = PDSALOWBC(NGBC(x))

SCOREBC = Round((1 / 8) * (((pgchs1 - pgchds1) + (pgphs1 - pgphds1) + (ngchs1 - ngchds1) + (ngphs1 - ngphds1)) - ((pgcls1 - pgclds1) + (pgpls1 - pgplds1) + (ngcls1 - ngclds1) + (ngpls1 - ngplds1))), 4)
End Function

Function ACCURACYBC(x As String) As Double
Dim i As Integer
'Positive grade
pgchs1 = CSAHIGHBC(PGBC(x))
pgphs1 = PSAHIGHBC(PGBC(x))
pgcls1 = CSALOWBC(PGBC(x))
pgpls1 = PSALOWBC(PGBC(x))
pgchds1 = CDSAHIGHBC(PGBC(x))
pgphds1 = PDSAHIGHBC(PGBC(x))
pgclds1 = CDSALOWBC(PGBC(x))
pgplds1 = PDSALOWBC(PGBC(x))
'Negative grade
ngchs1 = CSAHIGHBC(NGBC(x))
ngphs1 = PSAHIGHBC(NGBC(x))
ngcls1 = CSALOWBC(NGBC(x))
ngpls1 = PSALOWBC(NGBC(x))
ngchds1 = CDSAHIGHBC(NGBC(x))
ngphds1 = PDSAHIGHBC(NGBC(x))
ngclds1 = CDSALOWBC(NGBC(x))
ngplds1 = PDSALOWBC(NGBC(x))

ACCURACYBC = Round((1 / 8) * (((pgchs1 + pgchds1) / 2 + (pgphs1 + pgphds1) / 2 + (ngchs1 + ngchds1) + (ngphs1 + ngphds1)) - ((pgcls1 + pgclds1) / 2 + (pgpls1 + pgplds1) / 2 + (ngcls1 - ngclds1) + (ngpls1 - ngplds1))), 4)
End Function

Function CHECKBC(x As String) As Integer
Dim dem As Integer
pgchs1 = CSAHIGHBC(PGBC(x))
pgphs1 = PSAHIGHBC(PGBC(x))
pgcls1 = CSALOWBC(PGBC(x))
pgpls1 = PSALOWBC(PGBC(x))
pgchds1 = CDSAHIGHBC(PGBC(x))
pgphds1 = PDSAHIGHBC(PGBC(x))
pgclds1 = CDSALOWBC(PGBC(x))
pgplds1 = PDSALOWBC(PGBC(x))
'Negative grade
ngchs1 = CSAHIGHBC(NGBC(x))
ngphs1 = PSAHIGHBC(NGBC(x))
ngcls1 = CSALOWBC(NGBC(x))
ngpls1 = PSALOWBC(NGBC(x))
ngchds1 = CDSAHIGHBC(NGBC(x))
ngphds1 = PDSAHIGHBC(NGBC(x))
ngclds1 = CDSALOWBC(NGBC(x))
ngplds1 = PDSALOWBC(NGBC(x))

If (pgchds1 + pgclds1 >= 0) Then
    If (pgchds1 + pgclds1 <= 1) Then
        dem = dem + 1
    End If
End If

If (ngchds1 + ngclds1 >= -1) Then
    If (ngchds1 + ngclds1 <= 0) Then
        dem = dem + 1
    End If
End If
    
If (pgchds1 * pgchs1 + pgclds1 * pgcls1 >= 0) Then
    If (pgchds1 * pgchs1 + pgclds1 * pgcls1 <= 1) Then
        dem = dem + 1
    End If
End If

If (ngchds1 * ngchs1 + ngclds1 * ngcls1 >= 0) Then
    If (ngchds1 * ngchs1 + ngclds1 * ngcls1 <= 1) Then
        dem = dem + 1
    End If
End If

If (pgphds1 + pgplds1 >= 0) Then
    If (pgphds1 + pgplds1 <= 1) Then
        dem = dem + 1
    End If
End If

If (ngphds1 + ngplds1 >= -1) Then
    If (ngphds1 + ngplds1 <= 0) Then
        dem = dem + 1
    End If
End If
    
If (pgphds1 * pgphs1 + pgplds1 * pgpls1 >= 0) Then
    If (pgphds1 * pgphs1 + pgplds1 * pgpls1 <= 1) Then
        dem = dem + 1
    End If
End If

If (ngphds1 * ngphs1 + ngplds1 * ngpls1 >= 0) Then
    If (ngphds1 * ngphs1 + ngplds1 * ngpls1 <= 1) Then
        dem = dem + 1
    End If
End If

CHECKBC = dem
End Function


