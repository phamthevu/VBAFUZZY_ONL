Sub CompleteIntitalMatrix()
Dim i As Integer, j As Integer, size As Integer
Dim FirstCol As Integer, FirstRow As Integer, startcol As Integer, startrow As Integer
Dim n As Integer, m As Integer

FirstCol = 2
FirstRow = 2
size = 15
startrow = FirstRow + size + 2
startcol = 2

For i = 1 To size
    For j = i To size
    
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "V" Then
            Cells(startrow + i - 1, startcol + j - 1) = 1
            Cells(startrow + j - 1, startcol + i - 1) = 0
        End If
        
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "A" Then
            Cells(startrow + i - 1, startcol + j - 1) = 0
            Cells(startrow + j - 1, startcol + i - 1) = 1
        End If
        
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "X" Then
            Cells(startrow + i - 1, startcol + j - 1) = 1
            Cells(startrow + j - 1, startcol + i - 1) = 1
        End If
        
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "O" Then
            Cells(startrow + i - 1, startcol + j - 1) = 0
            Cells(startrow + j - 1, startcol + i - 1) = 0
        End If
    
    Next j
Next i
End Sub

Sub CompleteIntitalMatrixWithFuzzy()
Dim i As Integer, j As Integer, size As Integer
Dim FirstCol As Integer, FirstRow As Integer, startcol As Integer, startrow As Integer
Dim FirstValueCol As Integer, FirstValueRow As Integer
Dim n As Integer, m As Integer

FirstCol = 2
FirstRow = 2
FirstValueRow = 2
FirstValueCol = 9
size = 5
startrow = FirstRow + size + 2
startcol = 2

For i = 1 To size
    For j = i To size
    
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "V" Then
            Cells(startrow + i - 1, startcol + j - 1) = Cells(FirstValueRow + i - 1, FirstValueCol + j - 1)
            Cells(startrow + j - 1, startcol + i - 1) = "NF"
        End If
        
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "A" Then
            Cells(startrow + i - 1, startcol + j - 1) = "NF"
            Cells(startrow + j - 1, startcol + i - 1) = Cells(FirstValueRow + i - 1, FirstValueCol + j - 1)
        End If
        
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "X" Then
            Cells(startrow + i - 1, startcol + j - 1) = Cells(FirstValueRow + i - 1, FirstValueCol + j - 1)
            Cells(startrow + j - 1, startcol + i - 1) = Cells(FirstValueRow + i - 1, FirstValueCol + j - 1)
        End If
        
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = "O" Then
            Cells(startrow + i - 1, startcol + j - 1) = "NF"
            Cells(startrow + j - 1, startcol + i - 1) = "NF"
        End If
    
    Next j
Next i
End Sub