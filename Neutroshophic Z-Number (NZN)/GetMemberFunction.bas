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
