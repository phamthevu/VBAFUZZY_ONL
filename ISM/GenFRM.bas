Sub GenFRM()
Dim FirstCol As Integer, FirstRow As Integer, startcol As Integer, startrow As Integer
Dim i As Integer, j As Integer, k As Integer, size As Integer
Dim check As Boolean
Dim nFRM As Integer, mFRM As Integer, nFR As Integer, mFR As Integer

FirstCol = 2
FirstRow = 18
size = 5
startrow = FirstRow + size + 2
startcol = FirstCol

For i = 1 To size
    For j = 1 To size
        check = False
        nFRM = startrow + i - 1
        mFRM = startcol + j - 1
        nRM = FirstRow + i - 1
        mRM = FirstCol + j - 1
        If Cells(nRM, mRM) <> 0 Then
            Cells(nFRM, mFRM) = Cells(nRM, mRM)
            check = True
        End If
        If Cells(nRM, mRM) = 0 Then
            For k = 1 To size
            If (Cells(FirstRow + k - 1, mRM) > 0) And (Cells(nRM, FirstCol + k - 1) > 0) Then
                Cells(nFRM, mFRM) = "1"
            check = True
            Exit For
            End If
        Next k
        End If
        If check = False Then Cells(nFRM, mFRM) = 0
    Next j
Next i
End Sub