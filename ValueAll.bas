Sub VALUEALL()
Dim i As Integer, n As Integer
Dim s As String
n = Worksheets.Count
For i = 1 To n
    Sheets(i).Activate
    Cells.Select
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False
    Application.CutCopyMode = False
    ActiveWorkbook.Save
    Range("K13").Select
Next i
End Sub