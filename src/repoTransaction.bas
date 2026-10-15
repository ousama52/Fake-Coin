Attribute VB_Name = "repoTransaction"
Option Explicit

' tblTransactions columns
Private Const C_ID As Long = 1
Private Const C_HEIGHT As Long = 2
Private Const C_POS As Long = 3
Private Const C_FROM As Long = 4
Private Const C_PUBKEY As Long = 5
Private Const C_TO As Long = 6
Private Const C_AMOUNT As Long = 7
Private Const C_FEE As Long = 8
Private Const C_NONCE As Long = 9
Private Const C_TIME As Long = 10
Private Const C_SIG As Long = 11
Private Const C_COUNT As Long = 11

' Every confirmed transaction in table order (clsTransaction objects).
Public Function LoadTransactions() As Collection
    Dim data As Variant
    Dim i As Long
    Dim list As Collection
    Dim tx As clsTransaction

    Set list = New Collection
    data = modSheetUtils.ReadTable(modSheetUtils.GetTable(modConstants.TBL_TRANSACTIONS))
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            Set tx = New clsTransaction
            tx.TxID = modSheetUtils.CellText(data(i, C_ID))
            tx.BlockHeight = modSheetUtils.CellLong(data(i, C_HEIGHT))
            tx.Position = modSheetUtils.CellLong(data(i, C_POS))
            tx.FromAddress = modSheetUtils.CellText(data(i, C_FROM))
            tx.FromPubKey = modSheetUtils.CellText(data(i, C_PUBKEY))
            tx.ToAddress = modSheetUtils.CellText(data(i, C_TO))
            tx.Amount = modSheetUtils.CellLong(data(i, C_AMOUNT))
            tx.Fee = modSheetUtils.CellLong(data(i, C_FEE))
            tx.Nonce = modSheetUtils.CellLong(data(i, C_NONCE))
            tx.Timestamp = modSheetUtils.CellLong(data(i, C_TIME))
            tx.Signature = modSheetUtils.CellText(data(i, C_SIG))
            list.Add tx
        Next i
    End If
    Set LoadTransactions = list
End Function

' Writes the transactions of every block in the chain (block order, then
' Position order) in one assignment. A chain with no transactions clears the table.
Public Sub SaveChainTransactions(ByVal chain As Collection)
    Dim arr() As Variant
    Dim data As Variant
    Dim b As Long
    Dim t As Long
    Dim n As Long
    Dim total As Long
    Dim blk As clsBlock
    Dim tx As clsTransaction

    For b = 1 To chain.Count
        total = total + chain(b).Transactions.Count
    Next b
    If total > 0 Then
        ReDim arr(1 To total, 1 To C_COUNT)
        For b = 1 To chain.Count
            Set blk = chain(b)
            For t = 1 To blk.Transactions.Count
                Set tx = blk.Transactions(t)
                n = n + 1
                arr(n, C_ID) = tx.TxID
                arr(n, C_HEIGHT) = blk.Height
                arr(n, C_POS) = t - 1
                arr(n, C_FROM) = tx.FromAddress
                arr(n, C_PUBKEY) = tx.FromPubKey
                arr(n, C_TO) = tx.ToAddress
                arr(n, C_AMOUNT) = tx.Amount
                arr(n, C_FEE) = tx.Fee
                arr(n, C_NONCE) = tx.Nonce
                arr(n, C_TIME) = tx.Timestamp
                arr(n, C_SIG) = tx.Signature
            Next t
        Next b
        data = arr
    End If
    modSheetUtils.WriteTable modSheetUtils.GetTable(modConstants.TBL_TRANSACTIONS), data
End Sub
