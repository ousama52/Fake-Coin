Attribute VB_Name = "repoBlock"
Option Explicit

' tblBlocks columns
Private Const C_HEIGHT As Long = 1
Private Const C_TIME As Long = 2
Private Const C_PREV As Long = 3
Private Const C_MERKLE As Long = 4
Private Const C_NONCE As Long = 5
Private Const C_BITS As Long = 6
Private Const C_MINER As Long = 7
Private Const C_TXCOUNT As Long = 8
Private Const C_HASH As Long = 9
Private Const C_WORK As Long = 10
Private Const C_SOURCE As Long = 11
Private Const C_COUNT As Long = 11

' The local best chain: every block in height order with its transactions
' attached (Position order). Index in the collection = Height + 1.
Public Function LoadChain() As Collection
    Dim data As Variant
    Dim txs As Collection
    Dim chain As Collection
    Dim byHeight As Object
    Dim blk As clsBlock
    Dim tx As clsTransaction
    Dim i As Long

    Set chain = New Collection
    data = modSheetUtils.ReadTable(modSheetUtils.GetTable(modConstants.TBL_BLOCKS))
    If IsEmpty(data) Then
        Set LoadChain = chain
        Exit Function
    End If
    For i = 1 To UBound(data, 1)
        chain.Add RowToBlock(data, i)
    Next i

    Set txs = repoTransaction.LoadTransactions()
    Set byHeight = CreateObject("Scripting.Dictionary")
    For i = 1 To txs.Count
        Set tx = txs(i)
        If Not byHeight.Exists(tx.BlockHeight) Then byHeight.Add tx.BlockHeight, New Collection
        InsertByPosition byHeight(tx.BlockHeight), tx
    Next i
    For i = 1 To chain.Count
        Set blk = chain(i)
        If byHeight.Exists(blk.Height) Then Set blk.Transactions = byHeight(blk.Height)
    Next i
    Set LoadChain = chain
End Function

' Writes Blocks and Transactions from the chain, each in one assignment.
' Callers set CumulativeWork first (svcDifficulty.AssignCumulativeWork).
Public Sub SaveChain(ByVal chain As Collection)
    Dim arr() As Variant
    Dim data As Variant
    Dim i As Long
    Dim blk As clsBlock

    If chain.Count > 0 Then
        ReDim arr(1 To chain.Count, 1 To C_COUNT)
        For i = 1 To chain.Count
            Set blk = chain(i)
            arr(i, C_HEIGHT) = blk.Height
            arr(i, C_TIME) = blk.Timestamp
            arr(i, C_PREV) = blk.PrevHash
            arr(i, C_MERKLE) = blk.MerkleRoot
            arr(i, C_NONCE) = blk.Nonce
            arr(i, C_BITS) = blk.DifficultyBits
            arr(i, C_MINER) = blk.Miner
            arr(i, C_TXCOUNT) = blk.TxCount
            arr(i, C_HASH) = blk.Hash
            arr(i, C_WORK) = blk.CumulativeWork
            arr(i, C_SOURCE) = blk.Source
        Next i
        data = arr
    End If
    modSheetUtils.WriteTable modSheetUtils.GetTable(modConstants.TBL_BLOCKS), data
    repoTransaction.SaveChainTransactions chain
End Sub

Private Function RowToBlock(ByRef data As Variant, ByVal r As Long) As clsBlock
    Dim b As clsBlock

    Set b = New clsBlock
    b.Height = modSheetUtils.CellLong(data(r, C_HEIGHT))
    b.Timestamp = modSheetUtils.CellLong(data(r, C_TIME))
    b.PrevHash = modSheetUtils.CellText(data(r, C_PREV))
    b.MerkleRoot = modSheetUtils.CellText(data(r, C_MERKLE))
    b.Nonce = modSheetUtils.CellLong(data(r, C_NONCE))
    b.DifficultyBits = modSheetUtils.CellLong(data(r, C_BITS))
    b.Miner = modSheetUtils.CellText(data(r, C_MINER))
    b.TxCount = modSheetUtils.CellLong(data(r, C_TXCOUNT))
    b.Hash = modSheetUtils.CellText(data(r, C_HASH))
    b.CumulativeWork = modSheetUtils.CellDecimal(data(r, C_WORK))
    b.Source = modSheetUtils.CellText(data(r, C_SOURCE))
    Set RowToBlock = b
End Function

' Keeps a block's transactions in Position order whatever the table order is.
Private Sub InsertByPosition(ByVal list As Collection, ByVal tx As clsTransaction)
    Dim j As Long
    Dim other As clsTransaction

    For j = 1 To list.Count
        Set other = list(j)
        If other.Position > tx.Position Then
            list.Add tx, , j
            Exit Sub
        End If
    Next j
    list.Add tx
End Sub
