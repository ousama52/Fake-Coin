Attribute VB_Name = "repoMempool"
Option Explicit

' tblMempool columns
Private Const C_ID As Long = 1
Private Const C_FROM As Long = 2
Private Const C_PUBKEY As Long = 3
Private Const C_TO As Long = 4
Private Const C_AMOUNT As Long = 5
Private Const C_FEE As Long = 6
Private Const C_NONCE As Long = 7
Private Const C_TIME As Long = 8
Private Const C_SIG As Long = 9
Private Const C_SOURCE As Long = 10
Private Const C_COUNT As Long = 10

' Pending transactions in table order (clsTransaction objects).
Public Function LoadMempool() As Collection
    Dim data As Variant
    Dim i As Long
    Dim list As Collection
    Dim tx As clsTransaction

    Set list = New Collection
    data = modSheetUtils.ReadTable(modSheetUtils.GetTable(modConstants.TBL_MEMPOOL))
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            Set tx = New clsTransaction
            tx.TxID = modSheetUtils.CellText(data(i, C_ID))
            tx.FromAddress = modSheetUtils.CellText(data(i, C_FROM))
            tx.FromPubKey = modSheetUtils.CellText(data(i, C_PUBKEY))
            tx.ToAddress = modSheetUtils.CellText(data(i, C_TO))
            tx.Amount = modSheetUtils.CellLong(data(i, C_AMOUNT))
            tx.Fee = modSheetUtils.CellLong(data(i, C_FEE))
            tx.Nonce = modSheetUtils.CellLong(data(i, C_NONCE))
            tx.Timestamp = modSheetUtils.CellLong(data(i, C_TIME))
            tx.Signature = modSheetUtils.CellText(data(i, C_SIG))
            tx.Source = modSheetUtils.CellText(data(i, C_SOURCE))
            list.Add tx
        Next i
    End If
    Set LoadMempool = list
End Function

' Replaces the whole Mempool with the given list (an empty list clears it).
Public Sub SaveMempool(ByVal list As Collection)
    Dim arr() As Variant
    Dim data As Variant
    Dim i As Long
    Dim tx As clsTransaction

    If list.Count > 0 Then
        ReDim arr(1 To list.Count, 1 To C_COUNT)
        For i = 1 To list.Count
            Set tx = list(i)
            arr(i, C_ID) = tx.TxID
            arr(i, C_FROM) = tx.FromAddress
            arr(i, C_PUBKEY) = tx.FromPubKey
            arr(i, C_TO) = tx.ToAddress
            arr(i, C_AMOUNT) = tx.Amount
            arr(i, C_FEE) = tx.Fee
            arr(i, C_NONCE) = tx.Nonce
            arr(i, C_TIME) = tx.Timestamp
            arr(i, C_SIG) = tx.Signature
            arr(i, C_SOURCE) = tx.Source
        Next i
        data = arr
    End If
    modSheetUtils.WriteTable modSheetUtils.GetTable(modConstants.TBL_MEMPOOL), data
End Sub
