Sub GenIRNA()
Dim FirstCol As Integer, FirstRow As Integer, startcol As Integer, startrow As Integer
Dim i As Integer, j As Integer, k As Integer, size As Integer
Dim opt As Variant, item As Variant
Dim check As Boolean, count As Integer, hasNewlevel As Boolean, check_item As Boolean
Dim nFRM As Integer, mFRM As Integer, nFR As Integer, mFR As Integer
Dim ReachSet() As Collection
Dim AntecSet() As Collection
Dim ReachSet_Temp() As Collection
Dim AntecSet_Temp() As Collection
Dim InterSet() As Collection
Dim SaveResult() As Collection
Dim Assigned() As Boolean, iSSame As Boolean
Dim level As Integer
Dim getAdd As Boolean
Dim msgResult As String, msgR As String, msgA As String

FirstCol = 2
FirstRow = 25
size = 5
startrow = 31
startcol = 2

ReDim ReachSet(1 To size)
ReDim AntecSet(1 To size)
ReDim ReachSet_Temp(1 To size)
ReDim AntecSet_Temp(1 To size)
ReDim InterSet(1 To size)
ReDim SaveResult(1 To size)
ReDim Assigned(1 To size)


For i = 1 To size
    Set ReachSet(i) = New Collection
    Set AntecSet(i) = New Collection
    For j = 1 To size
        If Cells(FirstRow + i - 1, FirstCol + j - 1) = 1 Then ReachSet(i).Add (j)
        If Cells(FirstRow + j - 1, FirstCol + i - 1) = 1 Then AntecSet(i).Add (j)
        Next j
        
        msgR = ""
        msgA = ""
        For Each opt In ReachSet(i)
            msgR = msgR & opt & ", "
        Next opt
        For Each opt In AntecSet(i)
            msgA = msgA & opt & ", "
        Next opt

        'Cells(startrow + i - 1, startcol).Value = "R(" & i & "): " & msgR
        'Cells(startrow + i - 1, startcol + 1).Value = "A(" & i & "): " & msgA
    Next i
   
For i = 1 To size
    Assigned(i) = False
Next i
count = 0
level = 0

Do
    hasNewlevel = False
    
    For i = 1 To size

        Set ReachSet_Temp(i) = New Collection
        Set AntecSet_Temp(i) = New Collection
        For j = 1 To ReachSet(i).count
            If Assigned(ReachSet(i)(j)) = False Then
                ReachSet_Temp(i).Add (ReachSet(i)(j))
            End If
        Next j
        For j = 1 To AntecSet(i).count
            If Assigned(AntecSet(i)(j)) = False Then
                AntecSet_Temp(i).Add (AntecSet(i)(j))
            End If
        Next j
        Set ReachSet(i) = ReachSet_Temp(i)
        Set ReachSet(i) = ReachSet_Temp(i)
 
        If Assigned(i) = False Then
            Set InterSet(i) = New Collection
            For Each item In ReachSet(i)
                For k = 1 To AntecSet(i).count
                    If item = AntecSet(i)(k) Then
                        InterSet(i).Add (item)
                        Exit For
                    End If
                Next k
            Next item
            
            iSSame = False
            If ReachSet(i).count = InterSet(i).count Then
                iSSame = True
                For j = 1 To ReachSet(i).count
                If ReachSet(i)(j) <> InterSet(i)(j) Then
                    iSSame = False
                    Exit For
                End If
            Next j
            End If
            
            If iSSame = True Then
                count = count + 1
                Assigned(i) = True
                If hasNewlevel = False Then
                    hasNewlevel = True
                    level = level + 1
                    Set SaveResult(level) = New Collection
                End If
                SaveResult(level).Add (i)
            End If
        End If
            
    Next i
    msgResult = ""
    For Each opt In SaveResult(level)
        msgResult = msgResult & opt & ","
    Next opt
    Cells(startrow + level, startcol - 1) = level
    Cells(startrow + level, startcol) = msgResult
Loop While count < size
End Sub