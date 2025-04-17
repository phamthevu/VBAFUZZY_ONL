Function NEITL(x As String) As Double
Dim i As Integer, j As Integer, start_pos As Integer, end_pos As Integer, mid_pos As Integer
Dim t As String, TL As String

start_pos = WorksheetFunction.Find("(", x, 1)
end_pos = WorksheetFunction.Find(")", x, 1)
t = Trim(Mid(x, start_pos, end_pos - start_pos + 1))

mid_pos = WorksheetFunction.Find(";", t, 1)
TL = Trim(Mid(t, 2, mid_pos - 2))

NEITL = CDbl(TL)
End Function

Function NEITU(x As String) As Double
Dim i As Integer, j As Integer, start_pos As Integer, end_pos As Integer, mid_pos As Integer
Dim t As String, TU As String

start_pos = WorksheetFunction.Find("(", x, 1)
end_pos = WorksheetFunction.Find(")", x, 1)
t = Trim(Mid(x, start_pos, end_pos - start_pos + 1))

mid_pos = WorksheetFunction.Find(";", t, 1)
TU = Trim(Mid(t, mid_pos + 1, end_pos - mid_pos - 2))

NEITU = CDbl(TU)
End Function

Function NEIIL(x As String) As Double
Dim start_pos As Integer, end_pos As Integer, mid_pos As Integer, prev_start_pos As Integer, prev_end_pos As Integer
Dim i As String, IL As String

prev_start_pos = WorksheetFunction.Find("(", x, 1)
prev_end_pos = WorksheetFunction.Find(")", x, 1)
start_pos = WorksheetFunction.Find("(", x, prev_start_pos + 1)
end_pos = WorksheetFunction.Find(")", x, prev_end_pos + 1)

i = Trim(Mid(x, start_pos, end_pos - start_pos + 1))

mid_pos = WorksheetFunction.Find(";", i, 1)
IL = Trim(Mid(i, 2, mid_pos - 2))

NEIIL = CDbl(IL)
End Function

Function NEIIU(x As String) As Double
Dim start_pos As Integer, end_pos As Integer, mid_pos As Integer, prev_start_pos As Integer, prev_end_pos As Integer
Dim i As String, IU As String

prev_start_pos = WorksheetFunction.Find("(", x, 1)
prev_end_pos = WorksheetFunction.Find(")", x, 1)
start_pos = WorksheetFunction.Find("(", x, prev_start_pos + 1)
end_pos = WorksheetFunction.Find(")", x, prev_end_pos + 1)

i = Trim(Mid(x, start_pos, end_pos - start_pos + 1))

mid_pos = WorksheetFunction.Find(";", i, 1)
IU = Trim(Mid(i, mid_pos + 1, Len(i) - mid_pos - 1))

NEIIU = CDbl(IU)
End Function

Function NEIFL(x As String) As Double
Dim start_pos As Integer, end_pos As Integer, mid_pos As Integer, prev_start_pos As Integer, prev_end_pos As Integer, first_prev_start_pos As Integer, first_prev_end_pos As Integer
Dim F As String, FL As String

first_prev_start_pos = WorksheetFunction.Find("(", x, 1)
first_prev_end_pos = WorksheetFunction.Find(")", x, 1)
prev_start_pos = WorksheetFunction.Find("(", x, first_prev_start_pos + 1)
prev_end_pos = WorksheetFunction.Find(")", x, first_prev_end_pos + 1)
start_pos = WorksheetFunction.Find("(", x, prev_start_pos + 1)
end_pos = WorksheetFunction.Find(")", x, prev_end_pos + 1)

F = Trim(Mid(x, start_pos, end_pos - start_pos + 1))

mid_pos = WorksheetFunction.Find(";", F, 1)
FL = Trim(Mid(F, 2, mid_pos - 2))

NEIFL = CDbl(FL)
End Function

Function NEIFU(x As String) As Double
Dim start_pos As Integer, end_pos As Integer, mid_pos As Integer, prev_start_pos As Integer, prev_end_pos As Integer, first_prev_start_pos As Integer, first_prev_end_pos As Integer
Dim F As String, FU As String

first_prev_start_pos = WorksheetFunction.Find("(", x, 1)
first_prev_end_pos = WorksheetFunction.Find(")", x, 1)
prev_start_pos = WorksheetFunction.Find("(", x, first_prev_start_pos + 1)
prev_end_pos = WorksheetFunction.Find(")", x, first_prev_end_pos + 1)
start_pos = WorksheetFunction.Find("(", x, prev_start_pos + 1)
end_pos = WorksheetFunction.Find(")", x, prev_end_pos + 1)

F = Trim(Mid(x, start_pos, end_pos - start_pos + 1))

mid_pos = WorksheetFunction.Find(";", F, 1)
FU = Trim(Mid(F, mid_pos + 1, Len(F) - mid_pos - 1))

NEIFU = CDbl(FU)
End Function




