Function QFSUC(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
t = Trim(Mid(s, 2, i - 2))
j = WorksheetFunction.Find("e", t, 1)
QFSUC = CDbl(Mid(t, 1, j - 1))
End Function

Function QFSUE(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
t = Trim(Mid(s, 2, i - 2))
j = WorksheetFunction.Find("e", t, 1)
QFSUE = CDbl(Mid(t, j + 5, Len(t) - j))
End Function

Function QFSVC(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, i + 2, j - i - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSVC = CDbl(Mid(t, 1, k - 1))
End Function

Function QFSVE(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, i + 2, j - i - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSVE = CDbl(Mid(t, k + 5, Len(t) - 1))
End Function
Function QFSHC(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, j + 2, Len(s) - j - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSHC = CDbl(Mid(t, 1, k - 1))
End Function

Function QFSHE(s As String) As Double
Dim i As Integer, t As String, j As Integer
i = WorksheetFunction.Find(";", s, 1)
j = WorksheetFunction.Find(";", s, WorksheetFunction.Find(";", s, 1) + 1)
t = Trim(Mid(s, j + 2, Len(s) - j - 2))
k = WorksheetFunction.Find("e", t, 1)
QFSHE = CDbl(Mid(t, k + 5, Len(t) - 1))
End Function












