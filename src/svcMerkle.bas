Attribute VB_Name = "svcMerkle"
Option Explicit

' Merkle root of a list of TxIDs (coinbase first), exactly as DATA_TABLES.md
' defines it:
'   - no transactions: SHA256("")
'   - otherwise pair neighbours with SHA256(left & right); an odd last item is
'     paired with itself; repeat until one hash is left.
' A single transaction therefore IS the root (there is nothing to pair).
Public Function MerkleRoot(ByVal txIds As Collection) As String
    Dim level() As String
    Dim nextLevel() As String
    Dim n As Long
    Dim m As Long
    Dim i As Long
    Dim leftHash As String
    Dim rightHash As String

    If txIds.Count = 0 Then
        MerkleRoot = svcHash.Sha256("")
        Exit Function
    End If
    n = txIds.Count
    ReDim level(1 To n)
    For i = 1 To n
        level(i) = CStr(txIds(i))
    Next i

    Do While n > 1
        m = (n + 1) \ 2
        ReDim nextLevel(1 To m)
        For i = 1 To m
            leftHash = level(2 * i - 1)
            If 2 * i <= n Then
                rightHash = level(2 * i)
            Else
                rightHash = leftHash
            End If
            nextLevel(i) = svcHash.Sha256(leftHash & rightHash)
        Next i
        level = nextLevel
        n = m
    Loop
    MerkleRoot = level(1)
End Function
