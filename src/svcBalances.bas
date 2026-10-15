Attribute VB_Name = "svcBalances"
Option Explicit

' Chain state = a 4-element array of Scripting.Dictionary objects:
'   state(0) address -> balance (everything received minus everything sent)
'   state(1) address -> next nonce (confirmed transactions sent)
'   state(2) TxID -> True (every confirmed transaction)
'   state(3) block height -> Array(address, amount) of that block's coinbase
' Replaying the chain into a fresh state is how every rule check, mining
' selection and balance rebuild gets its numbers.

Public Function NewState() As Variant
    Dim st(0 To 3) As Variant
    Dim i As Long

    For i = 0 To 3
        Set st(i) = CreateObject("Scripting.Dictionary")
    Next i
    NewState = st
End Function

Public Function BalanceOf(ByRef state As Variant, ByVal address As String) As Long
    If state(0).Exists(address) Then BalanceOf = state(0).Item(address)
End Function

Public Function NextNonceOf(ByRef state As Variant, ByVal address As String) As Long
    If state(1).Exists(address) Then NextNonceOf = state(1).Item(address)
End Function

' Balance minus the coinbase coins that are still immature when a block at
' atHeight is checked. Coinbase coins of block H are spendable in a block at
' height B only when B - H >= maturity, so the immature ones are the coinbases
' of heights H with B - maturity < H <= B. Maturity 0 turns this off.
Public Function SpendableOf(ByRef state As Variant, ByVal address As String, ByVal atHeight As Long, _
                            ByVal maturity As Long) As Long
    Dim h As Long
    Dim cb As Variant

    SpendableOf = BalanceOf(state, address)
    If maturity <= 0 Then Exit Function
    For h = atHeight - maturity + 1 To atHeight
        If state(3).Exists(h) Then
            cb = state(3).Item(h)
            If cb(0) = address Then SpendableOf = SpendableOf - cb(1)
        End If
    Next h
End Function

' Applies one confirmed transaction. The sender pays Amount + Fee; the fee
' reaches the miner through the coinbase Amount (reward + fees), so coins are
' neither created nor lost here. The coinbase has no sender; its Nonce is the
' block height.
Public Sub ApplyTx(ByVal tx As clsTransaction, ByRef state As Variant)
    If tx.FromAddress = modConstants.COINBASE_KEY Then
        state(3).Item(tx.Nonce) = Array(tx.ToAddress, tx.Amount)
    Else
        state(0).Item(tx.FromAddress) = BalanceOf(state, tx.FromAddress) - tx.Amount - tx.Fee
        state(1).Item(tx.FromAddress) = NextNonceOf(state, tx.FromAddress) + 1
    End If
    state(0).Item(tx.ToAddress) = BalanceOf(state, tx.ToAddress) + tx.Amount
    state(2).Item(tx.TxID) = True
End Sub

' Replays every transaction of the chain without checking rules.
Public Function StateFromChain(ByVal chain As Collection) As Variant
    Dim st As Variant
    Dim b As Long
    Dim t As Long
    Dim blk As clsBlock

    st = NewState()
    For b = 1 To chain.Count
        Set blk = chain(b)
        For t = 1 To blk.Transactions.Count
            ApplyTx blk.Transactions(t), st
        Next t
    Next b
    StateFromChain = st
End Function

' Rewrites tblBalances from the chain: every wallet of this node first (even
' with 0), then every other address that ever appeared on the chain.
' Spendable is as of the next block (tip height + 1).
Public Sub RebuildBalances(ByVal chain As Collection)
    Dim st As Variant
    Dim params As Variant
    Dim wallets As Variant
    Dim rows() As Variant
    Dim data As Variant
    Dim k As Variant
    Dim addr As String
    Dim nextHeight As Long
    Dim i As Long
    Dim n As Long

    st = StateFromChain(chain)
    params = repoConfig.LoadChainParams()
    nextHeight = chain.Count
    wallets = repoWallet.LoadWallets()
    ReDim rows(1 To modSheetUtils.RowCount(wallets) + st(0).Count + 1, 1 To repoBalance.BAC_COUNT)

    For i = 1 To modSheetUtils.RowCount(wallets)
        addr = modSheetUtils.CellText(wallets(i, repoWallet.WLC_ADDRESS))
        n = n + 1
        FillRow rows, n, st, addr, modSheetUtils.CellText(wallets(i, repoWallet.WLC_NAME)), nextHeight, params
    Next i
    For Each k In st(0).Keys
        If repoWallet.FindWalletByAddress(wallets, CStr(k)) = 0 Then
            n = n + 1
            FillRow rows, n, st, CStr(k), "", nextHeight, params
        End If
    Next k
    data = modSheetUtils.TakeRows(rows, n)
    repoBalance.SaveBalances data
End Sub

Private Sub FillRow(ByRef rows() As Variant, ByVal n As Long, ByRef st As Variant, ByVal addr As String, _
                    ByVal walletName As String, ByVal nextHeight As Long, ByRef params As Variant)
    rows(n, repoBalance.BAC_ADDRESS) = addr
    rows(n, repoBalance.BAC_NAME) = walletName
    rows(n, repoBalance.BAC_BALANCE) = BalanceOf(st, addr)
    rows(n, repoBalance.BAC_SPENDABLE) = SpendableOf(st, addr, nextHeight, params(modConstants.PARAM_MATURITY))
    rows(n, repoBalance.BAC_NONCE) = NextNonceOf(st, addr)
End Sub
