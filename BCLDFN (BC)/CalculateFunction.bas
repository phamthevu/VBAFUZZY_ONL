Function ADDBC(x1 As String, x2 As String) As String
Dim pgchs1 As Double, pgphs1 As Double, pgcls1 As Double, pgpls1 As Double, pgchds1 As Double, pgphds1 As Double, pgclds1 As Double, pgplds1 As Double
Dim pgchs2 As Double, pgphs2 As Double, pgcls2 As Double, pgpls2 As Double, pgchds2 As Double, pgphds2 As Double, pgclds2 As Double, pgplds2 As Double
Dim pgchs As Double, pgphs As Double, pgcls As Double, pgpls As Double, pgchds As Double, pgphds As Double, pgclds As Double, pgplds As Double
Dim ngchs1 As Double, ngphs1 As Double, ngcls1 As Double, ngpls1 As Double, ngchds1 As Double, ngphds1 As Double, ngclds1 As Double, ngplds1 As Double
Dim ngchs2 As Double, ngphs2 As Double, ngcls2 As Double, ngpls2 As Double, ngchds2 As Double, ngphds2 As Double, ngclds2 As Double, ngplds2 As Double
Dim ngchs As Double, ngphs As Double, ngcls As Double, ngpls As Double, ngchds As Double, ngphds As Double, ngclds As Double, ngplds As Double
Dim pg As String, ng As String

'Positive grade
pgchs1 = CSAHIGHBC(PGBC(x1))
pgphs1 = PSAHIGHBC(PGBC(x1))
pgcls1 = CSALOWBC(PGBC(x1))
pgpls1 = PSALOWBC(PGBC(x1))
pgchds1 = CDSAHIGHBC(PGBC(x1))
pgphds1 = PDSAHIGHBC(PGBC(x1))
pgclds1 = CDSALOWBC(PGBC(x1))
pgplds1 = PDSALOWBC(PGBC(x1))

pgchs2 = CSAHIGHBC(PGBC(x2))
pgphs2 = PSAHIGHBC(PGBC(x2))
pgcls2 = CSALOWBC(PGBC(x2))
pgpls2 = PSALOWBC(PGBC(x2))
pgchds2 = CDSAHIGHBC(PGBC(x2))
pgphds2 = PDSAHIGHBC(PGBC(x2))
pgclds2 = CDSALOWBC(PGBC(x2))
pgplds2 = PDSALOWBC(PGBC(x2))
 
pgchs = Round(pgchs1 + pgchs2 - pgchs1 * pgchs2, 4)
pgphs = Round(pgphs1 + pgphs2 - pgphs1 * pgphs2, 4)
pgcls = Round(-(-pgcls1) * (-pgcls2), 4)
pgpls = Round(-(-pgpls1) * (-pgpls2), 4)
pgchds = Round(pgchds1 * pgchds2, 4)
pgphds = Round(pgphds1 * pgphds2, 4)
pgclds = Round(-((-pgclds1) + (-pgclds2) - (-pgclds1) * (-pgclds2)), 4)
pgplds = Round(-((-pgplds1) + (-pgplds2) - (-pgplds1) * (-pgplds2)), 4)

pg = "[" & pgchs & "ei2pi(" & pgphs & "); " & pgcls & "ei2pi(" & pgpls & "); " & pgchds & "ei2pi(" & pgphds & "); " & pgclds & "ei2pi( " & pgplds & ")]"

'Negative grade
ngchs1 = CSAHIGHBC(NGBC(x1))
ngphs1 = PSAHIGHBC(NGBC(x1))
ngcls1 = CSALOWBC(NGBC(x1))
ngpls1 = PSALOWBC(NGBC(x1))
ngchds1 = CDSAHIGHBC(NGBC(x1))
ngphds1 = PDSAHIGHBC(NGBC(x1))
ngclds1 = CDSALOWBC(NGBC(x1))
ngplds1 = PDSALOWBC(NGBC(x1))

ngchs2 = CSAHIGHBC(NGBC(x2))
ngphs2 = PSAHIGHBC(NGBC(x2))
ngcls2 = CSALOWBC(NGBC(x2))
ngpls2 = PSALOWBC(NGBC(x2))
ngchds2 = CDSAHIGHBC(NGBC(x2))
ngphds2 = PDSAHIGHBC(NGBC(x2))
ngclds2 = CDSALOWBC(NGBC(x2))
ngplds2 = PDSALOWBC(NGBC(x2))
 
ngchs = Round(ngchs1 + ngchs2 - ngchs1 * ngchs2, 4)
ngphs = Round(ngphs1 + ngphs2 - ngphs1 * ngphs2, 4)
ngcls = Round(-(-ngcls1) * (-ngcls2), 4)
ngpls = Round(-(-ngpls1) * (-ngpls2), 4)
ngchds = Round(ngchds1 * ngchds2, 4)
ngphds = Round(ngphds1 * ngphds2, 4)
ngclds = Round(-((-ngclds1) + (-ngclds2) - (-ngclds1) * (-ngclds2)), 4)
ngplds = Round(-((-ngplds1) + (-ngplds2) - (-ngplds1) * (-ngplds2)), 4)

ng = "[" & ngchs & "ei2pi(" & ngphs & "); " & ngcls & "ei2pi(" & ngpls & "); " & ngchds & "ei2pi(" & ngphds & "); " & ngclds & "ei2pi( " & ngplds & ")]"

ADDBC = "(" + pg + "; " + ng + ")"
End Function

Function MULBC(x1 As String, x2 As String) As String
Dim pgchs1 As Double, pgphs1 As Double, pgcls1 As Double, pgpls1 As Double, pgchds1 As Double, pgphds1 As Double, pgclds1 As Double, pgplds1 As Double
Dim pgchs2 As Double, pgphs2 As Double, pgcls2 As Double, pgpls2 As Double, pgchds2 As Double, pgphds2 As Double, pgclds2 As Double, pgplds2 As Double
Dim pgchs As Double, pgphs As Double, pgcls As Double, pgpls As Double, pgchds As Double, pgphds As Double, pgclds As Double, pgplds As Double
Dim ngchs1 As Double, ngphs1 As Double, ngcls1 As Double, ngpls1 As Double, ngchds1 As Double, ngphds1 As Double, ngclds1 As Double, ngplds1 As Double
Dim ngchs2 As Double, ngphs2 As Double, ngcls2 As Double, ngpls2 As Double, ngchds2 As Double, ngphds2 As Double, ngclds2 As Double, ngplds2 As Double
Dim ngchs As Double, ngphs As Double, ngcls As Double, ngpls As Double, ngchds As Double, ngphds As Double, ngclds As Double, ngplds As Double
Dim pg As String, ng As String

'Positive grade
pgchs1 = CSAHIGHBC(PGBC(x1))
pgphs1 = PSAHIGHBC(PGBC(x1))
pgcls1 = CSALOWBC(PGBC(x1))
pgpls1 = PSALOWBC(PGBC(x1))
pgchds1 = CDSAHIGHBC(PGBC(x1))
pgphds1 = PDSAHIGHBC(PGBC(x1))
pgclds1 = CDSALOWBC(PGBC(x1))
pgplds1 = PDSALOWBC(PGBC(x1))

pgchs2 = CSAHIGHBC(PGBC(x2))
pgphs2 = PSAHIGHBC(PGBC(x2))
pgcls2 = CSALOWBC(PGBC(x2))
pgpls2 = PSALOWBC(PGBC(x2))
pgchds2 = CDSAHIGHBC(PGBC(x2))
pgphds2 = PDSAHIGHBC(PGBC(x2))
pgclds2 = CDSALOWBC(PGBC(x2))
pgplds2 = PDSALOWBC(PGBC(x2))
 
pgchs = Round(pgchs1 * pgchs2, 4)
pgphs = Round(pgphs1 * pgphs2, 4)
pgcls = Round(-((-pgcls1) + (-pgcls2) - (-pgcls1) * (-pgcls2)), 4)
pgpls = Round(-((-pgpls1) + (-pgpls2) - (-pgpls1) * (-pgpls2)), 4)
pgchds = Round(pgchds1 + pgchds2 - pgchds1 * pgchds2, 4)
pgphds = Round(pgphds1 + pgphds2 - pgphds1 * pgphds2, 4)
pgclds = Round(-(-pgclds1) * (-pgclds2), 4)
pgplds = Round(-(-pgplds1) * (-pgplds2), 4)

pg = "[" & pgchs & "ei2pi(" & pgphs & "); " & pgcls & "ei2pi(" & pgpls & "); " & pgchds & "ei2pi(" & pgphds & "); " & pgclds & "ei2pi( " & pgplds & ")]"

'Negative grade
ngchs1 = CSAHIGHBC(NGBC(x1))
ngphs1 = PSAHIGHBC(NGBC(x1))
ngcls1 = CSALOWBC(NGBC(x1))
ngpls1 = PSALOWBC(NGBC(x1))
ngchds1 = CDSAHIGHBC(NGBC(x1))
ngphds1 = PDSAHIGHBC(NGBC(x1))
ngclds1 = CDSALOWBC(NGBC(x1))
ngplds1 = PDSALOWBC(NGBC(x1))

ngchs2 = CSAHIGHBC(NGBC(x2))
ngphs2 = PSAHIGHBC(NGBC(x2))
ngcls2 = CSALOWBC(NGBC(x2))
ngpls2 = PSALOWBC(NGBC(x2))
ngchds2 = CDSAHIGHBC(NGBC(x2))
ngphds2 = PDSAHIGHBC(NGBC(x2))
ngclds2 = CDSALOWBC(NGBC(x2))
ngplds2 = PDSALOWBC(NGBC(x2))
 
ngchs = Round(ngchs1 * ngchs2, 4)
ngphs = Round(ngphs1 * ngphs2, 4)
ngcls = Round(-((-ngcls1) + (-ngcls2) - (-ngcls1) * (-ngcls2)), 4)
ngpls = Round(-((-ngpls1) + (-ngpls2) - (-ngpls1) * (-ngpls2)), 4)
ngchds = Round(ngchds1 + ngchds2 - ngchds1 * ngchds2, 4)
ngphds = Round(ngphds1 + ngphds2 - ngphds1 * ngphds2, 4)
ngclds = Round(-(-ngclds1) * (-ngclds2), 4)
ngplds = Round(-(-ngplds1) * (-ngplds2), 4)

ng = "[" & ngchs & "ei2pi(" & ngphs & "); " & ngcls & "ei2pi(" & ngpls & "); " & ngchds & "ei2pi(" & ngphds & "); " & ngclds & "ei2pi( " & ngplds & ")]"
MULBC = "(" + pg + "; " + ng + ")"
End Function

Function COEBC(x As String, k As Double) As String
Dim pgchs1 As Double, pgphs1 As Double, pgcls1 As Double, pgpls1 As Double, pgchds1 As Double, pgphds1 As Double, pgclds1 As Double, pgplds1 As Double
Dim pgchs2 As Double, pgphs2 As Double, pgcls2 As Double, pgpls2 As Double, pgchds2 As Double, pgphds2 As Double, pgclds2 As Double, pgplds2 As Double
Dim pgchs As Double, pgphs As Double, pgcls As Double, pgpls As Double, pgchds As Double, pgphds As Double, pgclds As Double, pgplds As Double
Dim ngchs1 As Double, ngphs1 As Double, ngcls1 As Double, ngpls1 As Double, ngchds1 As Double, ngphds1 As Double, ngclds1 As Double, ngplds1 As Double
Dim ngchs2 As Double, ngphs2 As Double, ngcls2 As Double, ngpls2 As Double, ngchds2 As Double, ngphds2 As Double, ngclds2 As Double, ngplds2 As Double
Dim ngchs As Double, ngphs As Double, ngcls As Double, ngpls As Double, ngchds As Double, ngphds As Double, ngclds As Double, ngplds As Double
Dim pg As String, ng As String

'Positive grade
pgchs1 = CSAHIGHBC(PGBC(x))
pgphs1 = PSAHIGHBC(PGBC(x))
pgcls1 = CSALOWBC(PGBC(x))
pgpls1 = PSALOWBC(PGBC(x))
pgchds1 = CDSAHIGHBC(PGBC(x))
pgphds1 = PDSAHIGHBC(PGBC(x))
pgclds1 = CDSALOWBC(PGBC(x))
pgplds1 = PDSALOWBC(PGBC(x))

 
pgchs = Round(1 - (1 - pgchs1) ^ (k), 4)
pgphs = Round(1 - (1 - pgphs1) ^ (k), 4)
pgcls = Round(-(-pgcls1) ^ (k), 4)
pgpls = Round(-(-pgpls1) ^ (k), 4)
pgchds = Round((pgchds1) ^ (k), 4)
pgphds = Round((pgphds1) ^ (k), 4)
pgclds = Round(-(1 - (1 - (-pgclds1)) ^ (k)), 4)
pgplds = Round(-(1 - (1 - (-pgplds1)) ^ (k)), 4)

pg = "[" & pgchs & "ei2pi(" & pgphs & "); " & pgcls & "ei2pi(" & pgpls & "); " & pgchds & "ei2pi(" & pgphds & "); " & pgclds & "ei2pi( " & pgplds & ")]"

'Negative grade
ngchs1 = CSAHIGHBC(NGBC(x))
ngphs1 = PSAHIGHBC(NGBC(x))
ngcls1 = CSALOWBC(NGBC(x))
ngpls1 = PSALOWBC(NGBC(x))
ngchds1 = CDSAHIGHBC(NGBC(x))
ngphds1 = PDSAHIGHBC(NGBC(x))
ngclds1 = CDSALOWBC(NGBC(x))
ngplds1 = PDSALOWBC(NGBC(x))

ngchs = Round(1 - (1 - ngchs1) ^ (k), 4)
ngphs = Round(1 - (1 - ngphs1) ^ (k), 4)
ngcls = Round(-(-ngcls1) ^ (k), 4)
ngpls = Round(-(-ngpls1) ^ (k), 4)
ngchds = Round((ngchds1) ^ (k), 4)
ngphds = Round((ngphds1) ^ (k), 4)
ngclds = Round(-(1 - (1 - (-ngclds1)) ^ (k)), 4)
ngplds = Round(-(1 - (1 - (-ngplds1)) ^ (k)), 4)
ng = "[" & ngchs & "ei2pi(" & ngphs & "); " & ngcls & "ei2pi(" & ngpls & "); " & ngchds & "ei2pi(" & ngphds & "); " & ngclds & "ei2pi( " & ngplds & ")]"

COEBC = "(" + pg + "; " + ng + ")"
End Function

Function POWBC(x As String, k As Double) As String
Dim pgchs1 As Double, pgphs1 As Double, pgcls1 As Double, pgpls1 As Double, pgchds1 As Double, pgphds1 As Double, pgclds1 As Double, pgplds1 As Double
Dim pgchs2 As Double, pgphs2 As Double, pgcls2 As Double, pgpls2 As Double, pgchds2 As Double, pgphds2 As Double, pgclds2 As Double, pgplds2 As Double
Dim pgchs As Double, pgphs As Double, pgcls As Double, pgpls As Double, pgchds As Double, pgphds As Double, pgclds As Double, pgplds As Double
Dim ngchs1 As Double, ngphs1 As Double, ngcls1 As Double, ngpls1 As Double, ngchds1 As Double, ngphds1 As Double, ngclds1 As Double, ngplds1 As Double
Dim ngchs2 As Double, ngphs2 As Double, ngcls2 As Double, ngpls2 As Double, ngchds2 As Double, ngphds2 As Double, ngclds2 As Double, ngplds2 As Double
Dim ngchs As Double, ngphs As Double, ngcls As Double, ngpls As Double, ngchds As Double, ngphds As Double, ngclds As Double, ngplds As Double
Dim pg As String, ng As String

'Positive grade
pgchs1 = CSAHIGHBC(PGBC(x))
pgphs1 = PSAHIGHBC(PGBC(x))
pgcls1 = CSALOWBC(PGBC(x))
pgpls1 = PSALOWBC(PGBC(x))
pgchds1 = CDSAHIGHBC(PGBC(x))
pgphds1 = PDSAHIGHBC(PGBC(x))
pgclds1 = CDSALOWBC(PGBC(x))
pgplds1 = PDSALOWBC(PGBC(x))

pgchs = Round(pgchs1 ^ (k), 4)
pgphs = Round(pgphs1 ^ (k), 4)
pgcls = Round(-(1 - (1 - (-pgcls1)) ^ (k)), 4)
pgpls = Round(-(1 - (1 - (-pgpls1)) ^ (k)), 4)
pgchds = Round(1 - ((1 - pgchds1) ^ (k)), 4)
pgphds = Round(1 - ((1 - pgphds1) ^ (k)), 4)
pgclds = Round(-(-pgclds1) ^ (k), 4)
pgplds = Round(-(-pgplds1) ^ (k), 4)
pg = "[" & pgchs & "ei2pi(" & pgphs & "); " & pgcls & "ei2pi(" & pgpls & "); " & pgchds & "ei2pi(" & pgphds & "); " & pgclds & "ei2pi( " & pgplds & ")]"

'Negative grade
ngchs1 = CSAHIGHBC(NGBC(x))
ngphs1 = PSAHIGHBC(NGBC(x))
ngcls1 = CSALOWBC(NGBC(x))
ngpls1 = PSALOWBC(NGBC(x))
ngchds1 = CDSAHIGHBC(NGBC(x))
ngphds1 = PDSAHIGHBC(NGBC(x))
ngclds1 = CDSALOWBC(NGBC(x))
ngplds1 = PDSALOWBC(NGBC(x))

ngchs = Round(ngchs1 ^ (k), 4)
ngphs = Round(ngphs1 ^ (k), 4)
ngcls = Round(-(1 - (1 - (-ngcls1)) ^ (k)), 4)
ngpls = Round(-(1 - (1 - (-ngpls1)) ^ (k)), 4)
ngchds = Round(1 - ((1 - ngchds1) ^ (k)), 4)
ngphds = Round(1 - ((1 - ngphds1) ^ (k)), 4)
ngclds = Round(-(-ngclds1) ^ (k), 4)
ngplds = Round(-(-ngplds1) ^ (k), 4)
ng = "[" & ngchs & "ei2pi(" & ngphs & "); " & ngcls & "ei2pi(" & ngpls & "); " & ngchds & "ei2pi(" & ngphds & "); " & ngclds & "ei2pi( " & ngplds & ")]"

POWBC = "(" + pg + "; " + ng + ")"
End Function