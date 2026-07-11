Function PRO_PFS_SUM(x1 As String, x2 As String) As String
Dim pfs_x1 As String, pfs_x2 As String, sum As String

pfs_x1 = PRO_CONVERT_TO_PFS(x1)
pfs_x2 = PRO_CONVERT_TO_PFS(x2)

sum = PFS_SUM(pfs_x1, pfs_x2)

PRO_PFS_SUM = sum
End Function

Function PRO_PFS_MUL(x1 As String, x2 As String) As String
Dim pfs_x1 As String, pfs_x2 As String, mul As String

pfs_x1 = PRO_CONVERT_TO_PFS(x1)
pfs_x2 = PRO_CONVERT_TO_PFS(x2)

mul = PFS_MUL(pfs_x1, pfs_x2)

PRO_PFS_MUL = mul
End Function

Function PRO_PFS_COE(x1 As String, n As Double) As String
Dim pfs_x1 As String, pfs_x2 As String, coe As String

pfs_x1 = PRO_CONVERT_TO_PFS(x1)

coe = PFS_COE(pfs_x1, n)

PRO_PFS_COE = coe
End Function

Function PRO_PFS_POW(x1 As String, n As Double) As String
Dim pfs_x1 As String, pfs_x2 As String, pow As String

pfs_x1 = PRO_CONVERT_TO_PFS(x1)

pow = PFS_POW(pfs_x1, n)

PRO_PFS_POW = pow
End Function