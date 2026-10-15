Attribute VB_Name = "svcTransaction"
Option Explicit

' TxID formula from DATA_TABLES.md. Integers are plain decimal text.
Public Function ComputeTxId(ByVal fromAddress As String, ByVal toAddress As String, ByVal amount As Long, _
                            ByVal fee As Long, ByVal nonce As Long, ByVal timestamp As Long) As String
    ComputeTxId = svcHash.Sha256(fromAddress & "|" & toAddress & "|" & CStr(amount) & "|" & CStr(fee) & "|" & _
                                 CStr(nonce) & "|" & CStr(timestamp))
End Function

' Builds a transaction signed with the sender's private key (RSA over the TxID text).
Public Function NewTransaction(ByVal privateKey As String, ByVal fromPubKey As String, ByVal fromAddress As String, _
                               ByVal toAddress As String, ByVal amount As Long, ByVal fee As Long, _
                               ByVal nonce As Long, ByVal timestamp As Long) As clsTransaction
    Dim tx As clsTransaction

    Set tx = New clsTransaction
    tx.FromAddress = fromAddress
    tx.FromPubKey = fromPubKey
    tx.ToAddress = toAddress
    tx.Amount = amount
    tx.Fee = fee
    tx.Nonce = nonce
    tx.Timestamp = timestamp
    tx.TxID = ComputeTxId(fromAddress, toAddress, amount, fee, nonce, timestamp)
    tx.Signature = svcCrypto.Sign(privateKey, tx.TxID)
    Set NewTransaction = tx
End Function

' Coinbase: no sender, no public key, no signature, Fee 0, Nonce = block height.
Public Function NewCoinbase(ByVal height As Long, ByVal minerAddress As String, ByVal amount As Long, _
                            ByVal timestamp As Long) As clsTransaction
    Dim tx As clsTransaction

    Set tx = New clsTransaction
    tx.FromAddress = modConstants.COINBASE_KEY
    tx.ToAddress = minerAddress
    tx.Amount = amount
    tx.Fee = 0
    tx.Nonce = height
    tx.Timestamp = timestamp
    tx.BlockHeight = height
    tx.Position = 0
    tx.TxID = ComputeTxId(tx.FromAddress, tx.ToAddress, amount, 0, height, timestamp)
    Set NewCoinbase = tx
End Function

' Rules that depend only on the transaction itself (rules 1 to 3).
Public Function CheckStatic(ByVal tx As clsTransaction, ByVal minFee As Long, ByRef reason As String) As Boolean
    If Len(tx.TxID) = 0 Or Len(tx.FromAddress) = 0 Or Len(tx.FromPubKey) = 0 _
       Or Len(tx.ToAddress) = 0 Or Len(tx.Signature) = 0 Then
        reason = modMessages.MSG_TX_R_FIELDS
    ElseIf tx.Amount < 1 Then
        reason = modMessages.MSG_TX_R_AMOUNT
    ElseIf tx.Fee < minFee Then
        reason = modMessages.FormatText(modMessages.MSG_TX_R_FEE, minFee)
    ElseIf tx.FromAddress = tx.ToAddress Then
        reason = modMessages.MSG_TX_R_SELF
    ElseIf Not svcWallet.IsAddress(tx.ToAddress) Then
        reason = modMessages.MSG_TX_R_TOADDR
    ElseIf tx.TxID <> ComputeTxId(tx.FromAddress, tx.ToAddress, tx.Amount, tx.Fee, tx.Nonce, tx.Timestamp) Then
        reason = modMessages.MSG_TX_R_ID
    ElseIf svcCrypto.AddressFromPublicKey(tx.FromPubKey) <> tx.FromAddress Then
        reason = modMessages.MSG_TX_R_ADDR
    ElseIf Not svcCrypto.VerifySignature(tx.FromPubKey, tx.TxID, tx.Signature) Then
        reason = modMessages.MSG_TX_R_SIG
    Else
        CheckStatic = True
    End If
End Function

' Rules that depend on the chain so far (rules 4 to 6): not already confirmed,
' the sender's next nonce, and enough SPENDABLE balance for a block at atHeight.
Public Function CheckAgainstState(ByVal tx As clsTransaction, ByRef state As Variant, ByVal atHeight As Long, _
                                  ByRef params As Variant, ByRef reason As String) As Boolean
    Dim expected As Long
    Dim available As Long

    expected = svcBalances.NextNonceOf(state, tx.FromAddress)
    available = svcBalances.SpendableOf(state, tx.FromAddress, atHeight, params(modConstants.PARAM_MATURITY))
    If state(2).Exists(tx.TxID) Then
        reason = modMessages.MSG_TX_R_DUP
    ElseIf tx.Nonce <> expected Then
        reason = modMessages.FormatText(modMessages.MSG_TX_R_NONCE, tx.Nonce, expected)
    ElseIf available < tx.Amount + tx.Fee Then
        reason = modMessages.FormatText(modMessages.MSG_TX_R_BALANCE, available, tx.Amount + tx.Fee)
    Else
        CheckAgainstState = True
    End If
End Function

' Consensus rules for one non-coinbase transaction in a block at atHeight.
Public Function CheckTransaction(ByVal tx As clsTransaction, ByRef params As Variant, ByRef state As Variant, _
                                 ByVal atHeight As Long, ByRef reason As String) As Boolean
    If Not CheckStatic(tx, params(modConstants.PARAM_MINFEE), reason) Then Exit Function
    CheckTransaction = CheckAgainstState(tx, state, atHeight, params, reason)
End Function

' Adds a transaction to the Mempool list after ALL transaction rules pass
' (signature included), counting this sender's pending transactions for the
' nonce and the balance. Returns ADM_OK, ADM_DUP, ADM_REJECT, or ADM_WAIT
' (the nonce is ahead: retry later). A full Mempool evicts the lowest-fee entry
' when the new fee is higher, otherwise rejects. state is the confirmed chain.
Public Function AdmitToMempool(ByVal tx As clsTransaction, ByVal mem As Collection, ByRef state As Variant, _
                               ByRef params As Variant, ByVal tipHeight As Long, ByRef reason As String) As Long
    Dim pendingCount As Long
    Dim pendingOut As Long
    Dim expected As Long
    Dim available As Long
    Dim lowest As Long

    If state(2).Exists(tx.TxID) Or InMempool(mem, tx.TxID) Then
        reason = modMessages.MSG_TX_R_DUP
        AdmitToMempool = modConstants.ADM_DUP
        Exit Function
    End If
    AdmitToMempool = modConstants.ADM_REJECT
    If Not CheckStatic(tx, params(modConstants.PARAM_MINFEE), reason) Then Exit Function

    PendingOf mem, tx.FromAddress, pendingCount, pendingOut
    expected = svcBalances.NextNonceOf(state, tx.FromAddress) + pendingCount
    If tx.Nonce > expected Then
        reason = modMessages.MSG_SYNC_NONCE_WAIT
        AdmitToMempool = modConstants.ADM_WAIT
        Exit Function
    End If
    If tx.Nonce < expected Then
        reason = modMessages.FormatText(modMessages.MSG_SYNC_NONCE_USED, tx.Nonce, expected)
        Exit Function
    End If
    available = svcBalances.SpendableOf(state, tx.FromAddress, tipHeight + 1, params(modConstants.PARAM_MATURITY)) - pendingOut
    If available < tx.Amount + tx.Fee Then
        reason = modMessages.FormatText(modMessages.MSG_TX_R_BALANCE, available, tx.Amount + tx.Fee)
        Exit Function
    End If

    If mem.Count >= params(modConstants.PARAM_MEMPOOL_MAX) Then
        lowest = LowestFeeIndex(mem)
        If lowest = 0 Then Exit Function
        If tx.Fee <= mem(lowest).Fee Then
            reason = modMessages.MSG_TX_R_MEMPOOL_FULL
            Exit Function
        End If
        mem.Remove lowest
    End If
    mem.Add tx
    AdmitToMempool = modConstants.ADM_OK
End Function

Private Function InMempool(ByVal mem As Collection, ByVal txId As String) As Boolean
    Dim i As Long

    For i = 1 To mem.Count
        If mem(i).TxID = txId Then
            InMempool = True
            Exit Function
        End If
    Next i
End Function

' How many Mempool transactions a sender has, and what they will spend.
Private Sub PendingOf(ByVal mem As Collection, ByVal address As String, ByRef count As Long, ByRef spend As Long)
    Dim i As Long

    count = 0
    spend = 0
    For i = 1 To mem.Count
        If mem(i).FromAddress = address Then
            count = count + 1
            spend = spend + mem(i).Amount + mem(i).Fee
        End If
    Next i
End Sub

' Position of the first lowest-fee entry, 0 for an empty list.
Private Function LowestFeeIndex(ByVal mem As Collection) As Long
    Dim i As Long

    For i = 1 To mem.Count
        If LowestFeeIndex = 0 Then
            LowestFeeIndex = i
        ElseIf mem(i).Fee < mem(LowestFeeIndex).Fee Then
            LowestFeeIndex = i
        End If
    Next i
End Function

' Signs a transaction with a wallet of this node and sends it: Mempool + tx_
' file. nonceText is the TEST-ONLY override (blank normally): with it the
' transaction is only written to the shared folder, not to this node's Mempool.
Public Function SendCoins(ByVal fromName As String, ByVal toInput As String, ByVal amountText As String, _
                          ByVal feeText As String, ByVal nonceText As String) As clsResult
    Dim params As Variant
    Dim st As Variant
    Dim mem As Collection
    Dim tx As clsTransaction
    Dim privateKey As String
    Dim fromAddr As String
    Dim fromPub As String
    Dim toAddr As String
    Dim reason As String
    Dim amount As Long
    Dim fee As Long
    Dim nonce As Long
    Dim chain As Collection
    Dim pendingCount As Long
    Dim pendingOut As Long

    params = repoConfig.LoadChainParams()
    privateKey = svcWallet.LoadPrivateKey(Trim$(fromName), fromAddr, fromPub, reason)
    If Len(privateKey) = 0 Then
        Set SendCoins = Fail(reason)
        Exit Function
    End If
    toAddr = svcWallet.ResolveAddress(toInput)
    If Len(toAddr) = 0 Then
        Set SendCoins = Fail(modMessages.FormatText(modMessages.MSG_WALLET_UNKNOWN, toInput))
        Exit Function
    End If
    If Not modSheetUtils.TryParseLong(amountText, amount) Or amount < 1 Then
        Set SendCoins = Fail(modMessages.MSG_TX_R_AMOUNT)
        Exit Function
    End If
    If Not modSheetUtils.TryParseLong(feeText, fee) Or fee < params(modConstants.PARAM_MINFEE) Then
        Set SendCoins = Fail(modMessages.FormatText(modMessages.MSG_TX_R_FEE, params(modConstants.PARAM_MINFEE)))
        Exit Function
    End If
    If fromAddr = toAddr Then
        Set SendCoins = Fail(modMessages.MSG_TX_R_SELF)
        Exit Function
    End If

    Set chain = repoBlock.LoadChain()
    st = svcBalances.StateFromChain(chain)
    Set mem = repoMempool.LoadMempool()
    PendingOf mem, fromAddr, pendingCount, pendingOut

    If Len(Trim$(nonceText)) > 0 Then
        If Not modSheetUtils.TryParseLong(nonceText, nonce) Then
            Set SendCoins = Fail(modMessages.MSG_TX_R_NONCE_INPUT)
            Exit Function
        End If
        Set SendCoins = SendNonceOverride(privateKey, fromPub, fromAddr, toAddr, amount, fee, nonce, st, chain.Count, params)
        Exit Function
    End If

    nonce = svcBalances.NextNonceOf(st, fromAddr) + pendingCount
    Set tx = NewTransaction(privateKey, fromPub, fromAddr, toAddr, amount, fee, nonce, modTime.UnixNowUtc())
    tx.Source = modConstants.SRC_LOCAL
    If AdmitToMempool(tx, mem, st, params, chain.Count - 1, reason) <> modConstants.ADM_OK Then
        Set SendCoins = Fail(reason)
        Exit Function
    End If

    svcNetwork.EnsureFolder
    repoMempool.SaveMempool mem
    svcNetwork.WriteTxFile tx
    clsLogger.LogInfo "svcTransaction.SendCoins", "Transaction " & tx.TxID & " added to the Mempool"
    Set SendCoins = Succeed(modMessages.FormatText(modMessages.MSG_TX_SENT, Left$(tx.TxID, 12), fromName, _
                            amount, toInput, fee, tx.Nonce))
End Function

' TEST ONLY: a signed transaction with a chosen nonce, written to the shared
' folder but not to this node's Mempool. It lets the double-spend test create
' two conflicting transactions with the same nonce.
Private Function SendNonceOverride(ByVal privateKey As String, ByVal fromPub As String, ByVal fromAddr As String, _
                                   ByVal toAddr As String, ByVal amount As Long, ByVal fee As Long, _
                                   ByVal nonce As Long, ByRef st As Variant, ByVal blockCount As Long, _
                                   ByRef params As Variant) As clsResult
    Dim tx As clsTransaction
    Dim reason As String
    Dim available As Long

    available = svcBalances.SpendableOf(st, fromAddr, blockCount, params(modConstants.PARAM_MATURITY))
    If available < amount + fee Then
        Set SendNonceOverride = Fail(modMessages.FormatText(modMessages.MSG_TX_R_BALANCE, available, amount + fee))
        Exit Function
    End If
    Set tx = NewTransaction(privateKey, fromPub, fromAddr, toAddr, amount, fee, nonce, modTime.UnixNowUtc())
    If Not CheckStatic(tx, params(modConstants.PARAM_MINFEE), reason) Then
        Set SendNonceOverride = Fail(reason)
        Exit Function
    End If
    svcNetwork.EnsureFolder
    svcNetwork.WriteTxFile tx
    clsLogger.LogWarning "svcTransaction.SendNonceOverride", "TEST ONLY: nonce override " & nonce & ", tx " & tx.TxID
    Set SendNonceOverride = Succeed(modMessages.FormatText(modMessages.MSG_TX_FILE_ONLY, nonce, Left$(tx.TxID, 12)))
End Function

' Pending (in the Mempool), Confirmed with N confirmations, or Unknown.
Public Function TxStatus(ByVal txId As String) As clsResult
    Dim mem As Collection
    Dim txs As Collection
    Dim tip As Long
    Dim i As Long
    Dim msg As String

    txId = Trim$(txId)
    If Len(txId) = 0 Then
        Set TxStatus = Fail(modMessages.MSG_TXSTATUS_NOINPUT)
        Exit Function
    End If
    msg = modMessages.FormatText(modMessages.MSG_TXSTATUS_UNKNOWN, Left$(txId, 12))
    Set mem = repoMempool.LoadMempool()
    For i = 1 To mem.Count
        If mem(i).TxID = txId Then msg = modMessages.FormatText(modMessages.MSG_TXSTATUS_PENDING, Left$(txId, 12))
    Next i
    Set txs = repoTransaction.LoadTransactions()
    tip = repoBlock.LoadChain().Count - 1
    For i = 1 To txs.Count
        If txs(i).TxID = txId Then
            msg = modMessages.FormatText(modMessages.MSG_TXSTATUS_CONFIRMED, Left$(txId, 12), tip - txs(i).BlockHeight + 1)
        End If
    Next i
    clsLogger.LogInfo "svcTransaction.TxStatus", msg
    Set TxStatus = Succeed(msg)
End Function

' ATTACK SIMULATION: a transaction that claims to be from the victim but is
' signed with a fresh random key. If the victim's public key is already on the
' chain it is used (bad signature); otherwise the attacker's own public key is
' used with the victim's address (address does not match). It is written as a
' normal tx_ file; other nodes must reject it at Sync.
Public Function ForgeTransaction(ByVal victimInput As String, ByVal toInput As String, ByVal amountText As String, _
                                 ByVal feeText As String) As clsResult
    Dim params As Variant
    Dim st As Variant
    Dim attackerKey As String
    Dim attackerPub As String
    Dim claimedPub As String
    Dim victim As String
    Dim toAddr As String
    Dim expected As String
    Dim amount As Long
    Dim fee As Long
    Dim tx As clsTransaction

    params = repoConfig.LoadChainParams()
    victim = svcWallet.ResolveAddress(victimInput)
    toAddr = svcWallet.ResolveAddress(toInput)
    If Len(victim) = 0 Or Len(toAddr) = 0 Then
        Set ForgeTransaction = Fail(modMessages.FormatText(modMessages.MSG_WALLET_UNKNOWN, IIf(Len(victim) = 0, victimInput, toInput)))
        Exit Function
    End If
    If Not modSheetUtils.TryParseLong(amountText, amount) Or amount < 1 Then
        Set ForgeTransaction = Fail(modMessages.MSG_TX_R_AMOUNT)
        Exit Function
    End If
    If Not modSheetUtils.TryParseLong(feeText, fee) Or fee < params(modConstants.PARAM_MINFEE) Then
        Set ForgeTransaction = Fail(modMessages.FormatText(modMessages.MSG_TX_R_FEE, params(modConstants.PARAM_MINFEE)))
        Exit Function
    End If

    svcCrypto.GenerateKeyPair attackerKey, attackerPub
    claimedPub = KnownPublicKey(victim)
    If Len(claimedPub) > 0 Then
        expected = modMessages.MSG_TX_R_SIG
    Else
        claimedPub = attackerPub
        expected = modMessages.MSG_TX_R_ADDR
    End If
    st = svcBalances.StateFromChain(repoBlock.LoadChain())
    Set tx = NewTransaction(attackerKey, claimedPub, victim, toAddr, amount, fee, _
                            svcBalances.NextNonceOf(st, victim), modTime.UnixNowUtc())
    svcNetwork.EnsureFolder
    svcNetwork.WriteTxFile tx
    clsLogger.LogWarning "svcTransaction.ForgeTransaction", "Forged " & tx.TxID & " as " & victim
    Set ForgeTransaction = Succeed(modMessages.FormatText(modMessages.MSG_FORGE_DONE, Left$(tx.TxID, 12), victim, expected))
End Function

' The public key of an address, if a confirmed transaction of theirs is on the chain.
Private Function KnownPublicKey(ByVal address As String) As String
    Dim txs As Collection
    Dim i As Long

    Set txs = repoTransaction.LoadTransactions()
    For i = 1 To txs.Count
        If txs(i).FromAddress = address And Len(txs(i).FromPubKey) > 0 Then
            KnownPublicKey = txs(i).FromPubKey
            Exit Function
        End If
    Next i
End Function

Private Function Fail(ByVal msg As String) As clsResult
    Dim r As clsResult

    Set r = New clsResult
    r.Init False, msg
    Set Fail = r
End Function

Private Function Succeed(ByVal msg As String) As clsResult
    Dim r As clsResult

    Set r = New clsResult
    r.Init True, msg, 1
    Set Succeed = r
End Function
