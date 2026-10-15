Attribute VB_Name = "svcValidation"
Option Explicit

' reward(height) = BlockReward \ 2 ^ ((height - 1) \ HalvingInterval), integer
' division. Halving one step at a time gives the same result without a
' Double power. Height 0 (genesis) has no reward.
Public Function BlockReward(ByVal height As Long, ByRef params As Variant) As Long
    Dim eras As Long
    Dim r As Long
    Dim i As Long

    If height < 1 Then Exit Function
    r = params(modConstants.PARAM_REWARD)
    eras = (height - 1) \ params(modConstants.PARAM_HALVING)
    For i = 1 To eras
        r = r \ 2
        If r = 0 Then Exit For
    Next i
    BlockReward = r
End Function

' Block hash formula from DATA_TABLES.md.
Public Function ComputeBlockHash(ByVal height As Long, ByVal timestamp As Long, ByVal prevHash As String, _
                                 ByVal merkle As String, ByVal nonce As Long, ByVal bits As Long, _
                                 ByVal miner As String) As String
    ComputeBlockHash = svcHash.Sha256(CStr(height) & "|" & CStr(timestamp) & "|" & prevHash & "|" & merkle & "|" & _
                                      CStr(nonce) & "|" & CStr(bits) & "|" & miner)
End Function

Public Function BlockHashOf(ByVal blk As clsBlock) As String
    BlockHashOf = ComputeBlockHash(blk.Height, blk.Timestamp, blk.PrevHash, blk.MerkleRoot, _
                                   blk.Nonce, blk.DifficultyBits, blk.Miner)
End Function

' The fixed genesis block, identical on every node.
Public Function GenesisBlock() As clsBlock
    Dim b As clsBlock

    Set b = New clsBlock
    b.Height = 0
    b.Timestamp = modConstants.GENESIS_TIMESTAMP
    b.PrevHash = modConstants.GENESIS_PREVHASH
    b.MerkleRoot = modConstants.GENESIS_MERKLE
    b.Nonce = 0
    b.DifficultyBits = modConstants.GENESIS_BITS
    b.Miner = modConstants.GENESIS_MINER
    b.TxCount = 0
    b.Hash = modConstants.GENESIS_HASH
    b.CumulativeWork = CDec(modConstants.GENESIS_WORK)
    b.Source = modConstants.SRC_GENESIS
    Set GenesisBlock = b
End Function

' Genesis is fixed: every field must equal the constants and the hash must
' recompute. It is exempt from the leading-zeros rule.
Public Function CheckGenesis(ByVal blk As clsBlock, ByRef reason As String) As Boolean
    Dim g As clsBlock

    Set g = GenesisBlock()
    If blk.Height = 0 And blk.Timestamp = g.Timestamp And blk.PrevHash = g.PrevHash _
       And blk.MerkleRoot = g.MerkleRoot And blk.Nonce = g.Nonce And blk.DifficultyBits = g.DifficultyBits _
       And blk.Miner = g.Miner And blk.TxCount = 0 And blk.Transactions.Count = 0 _
       And blk.Hash = g.Hash And BlockHashOf(blk) = g.Hash Then
        CheckGenesis = True
    Else
        reason = modMessages.MSG_BLK_GENESIS
    End If
End Function

' All block rules for one non-genesis block. chain holds the blocks before it
' and chain(parentIdx) is its parent (index = Height + 1). On success the
' block's transactions are applied to state; on failure state may be half
' updated, so callers must throw it away (they always replay into a fresh state).
Public Function CheckBlock(ByVal blk As clsBlock, ByVal chain As Collection, ByVal parentIdx As Long, _
                           ByRef params As Variant, ByRef state As Variant, ByRef reason As String) As Boolean
    If Not CheckHeader(blk, chain, parentIdx, params, reason) Then Exit Function
    CheckBlock = CheckBody(blk, params, state, reason)
End Function

' Height and parent link, the required difficulty bits, the hash and its
' proof of work, and the timestamp rules.
Private Function CheckHeader(ByVal blk As clsBlock, ByVal chain As Collection, ByVal parentIdx As Long, _
                             ByRef params As Variant, ByRef reason As String) As Boolean
    Dim parent As clsBlock
    Dim required As Long

    Set parent = chain(parentIdx)
    If blk.Height <> parent.Height + 1 Then
        reason = modMessages.FormatText(modMessages.MSG_BLK_HEIGHT, blk.Height)
        Exit Function
    End If
    If blk.PrevHash <> parent.Hash Then
        reason = modMessages.MSG_BLK_PREV
        Exit Function
    End If
    required = svcDifficulty.RequiredBits(blk.Height, chain, params)
    If blk.DifficultyBits <> required Then
        reason = modMessages.FormatText(modMessages.MSG_BLK_BITS, blk.DifficultyBits, required)
    ElseIf blk.Hash <> BlockHashOf(blk) Then
        reason = modMessages.MSG_BLK_HASH
    ElseIf svcHash.LeadingZeroBits(blk.Hash) < blk.DifficultyBits Then
        reason = modMessages.FormatText(modMessages.MSG_BLK_POW, svcHash.LeadingZeroBits(blk.Hash), blk.DifficultyBits)
    ElseIf blk.Timestamp <= parent.Timestamp Then
        reason = modMessages.MSG_BLK_TS_PARENT
    ElseIf blk.Timestamp > modTime.UnixNowUtc() + params(modConstants.PARAM_MAX_FUTURE) Then
        reason = modMessages.MSG_BLK_TS_FUTURE
    Else
        CheckHeader = True
    End If
End Function

' Transaction count, Merkle root, the single coinbase, its amount, the size
' limit and every other transaction in order.
Private Function CheckBody(ByVal blk As clsBlock, ByRef params As Variant, ByRef state As Variant, _
                           ByRef reason As String) As Boolean
    Dim txs As Collection
    Dim cb As clsTransaction
    Dim feeSum As Long
    Dim expected As Long

    Set txs = blk.Transactions
    If blk.TxCount <> txs.Count Then
        reason = modMessages.FormatText(modMessages.MSG_BLK_TXCOUNT, blk.TxCount, txs.Count)
        Exit Function
    End If
    If txs.Count = 0 Then
        reason = modMessages.MSG_BLK_NOCB
        Exit Function
    End If
    If blk.MerkleRoot <> svcMerkle.MerkleRoot(TxIdsOf(txs)) Then
        reason = modMessages.MSG_BLK_MERKLE
        Exit Function
    End If
    Set cb = txs(1)
    If cb.FromAddress <> modConstants.COINBASE_KEY Then
        reason = modMessages.MSG_BLK_NOCB
        Exit Function
    End If
    If Not SumFees(txs, params, feeSum, reason) Then Exit Function

    expected = BlockReward(blk.Height, params) + feeSum
    If Not CheckCoinbase(cb, blk.Height, expected, reason) Then Exit Function
    svcBalances.ApplyTx cb, state
    CheckBody = CheckOthers(txs, blk.Height, params, state, reason)
End Function

Private Function TxIdsOf(ByVal txs As Collection) As Collection
    Dim ids As Collection
    Dim i As Long

    Set ids = New Collection
    For i = 1 To txs.Count
        ids.Add txs(i).TxID
    Next i
    Set TxIdsOf = ids
End Function

' Fees of the non-coinbase transactions; also enforces "only one coinbase"
' and the MaxTxPerBlock limit.
Private Function SumFees(ByVal txs As Collection, ByRef params As Variant, ByRef feeSum As Long, _
                         ByRef reason As String) As Boolean
    Dim i As Long

    If txs.Count - 1 > params(modConstants.PARAM_MAXTX) Then
        reason = modMessages.FormatText(modMessages.MSG_BLK_MAXTX, txs.Count - 1, params(modConstants.PARAM_MAXTX))
        Exit Function
    End If
    For i = 2 To txs.Count
        If txs(i).FromAddress = modConstants.COINBASE_KEY Then
            reason = modMessages.MSG_BLK_MULTICB
            Exit Function
        End If
        feeSum = feeSum + txs(i).Fee
    Next i
    SumFees = True
End Function

Private Function CheckCoinbase(ByVal cb As clsTransaction, ByVal height As Long, ByVal expected As Long, _
                               ByRef reason As String) As Boolean
    If Len(cb.FromPubKey) > 0 Or Len(cb.Signature) > 0 Then
        reason = modMessages.MSG_BLK_CBFIELDS
    ElseIf cb.Nonce <> height Then
        reason = modMessages.MSG_BLK_CBNONCE
    ElseIf cb.Fee <> 0 Then
        reason = modMessages.MSG_BLK_CBFEE
    ElseIf cb.TxID <> svcTransaction.ComputeTxId(cb.FromAddress, cb.ToAddress, cb.Amount, cb.Fee, cb.Nonce, cb.Timestamp) Then
        reason = modMessages.MSG_BLK_CBID
    ElseIf cb.Amount <> expected Then
        reason = modMessages.FormatText(modMessages.MSG_BLK_CBAMOUNT, cb.Amount, expected)
    Else
        CheckCoinbase = True
    End If
End Function

' Every transaction after the coinbase, in block order, against the spendable
' balance at this block's height.
Private Function CheckOthers(ByVal txs As Collection, ByVal height As Long, ByRef params As Variant, _
                             ByRef state As Variant, ByRef reason As String) As Boolean
    Dim i As Long
    Dim why As String
    Dim tx As clsTransaction

    For i = 2 To txs.Count
        Set tx = txs(i)
        If Not svcTransaction.CheckTransaction(tx, params, state, height, why) Then
            reason = modMessages.FormatText(modMessages.MSG_BLK_TX, Left$(tx.TxID, 12), why)
            Exit Function
        End If
        svcBalances.ApplyTx tx, state
    Next i
    CheckOthers = True
End Function

' Replays a whole chain from genesis. badHeight is the position of the first
' block that breaks a rule (0 for the genesis block).
Public Function ReplayChain(ByVal chain As Collection, ByRef params As Variant, ByRef state As Variant, _
                            ByRef badHeight As Long, ByRef reason As String) As Boolean
    Dim i As Long

    state = svcBalances.NewState()
    If chain.Count = 0 Then
        reason = modMessages.MSG_BLK_EMPTYCHAIN
        Exit Function
    End If
    If Not CheckGenesis(chain(1), reason) Then Exit Function
    For i = 2 To chain.Count
        If Not CheckBlock(chain(i), chain, i - 1, params, state, reason) Then
            badHeight = i - 1
            Exit Function
        End If
    Next i
    ReplayChain = True
End Function

' Full check of every block from genesis (difficulty, timestamps, signatures
' and all other rules), reporting the first bad height.
Public Function ValidateChain() As clsResult
    Dim chain As Collection
    Dim params As Variant
    Dim state As Variant
    Dim badHeight As Long
    Dim reason As String
    Dim result As clsResult

    Set chain = repoBlock.LoadChain()
    params = repoConfig.LoadChainParams()
    Set result = New clsResult
    If ReplayChain(chain, params, state, badHeight, reason) Then
        result.Init True, modMessages.FormatText(modMessages.MSG_CHAIN_OK, chain.Count, chain.Count - 1, _
                          svcDifficulty.ChainWork(chain)), chain.Count
        clsLogger.LogInfo "svcValidation.ValidateChain", result.Message
    Else
        result.Init False, modMessages.FormatText(modMessages.MSG_CHAIN_BAD, badHeight, reason), badHeight
        clsLogger.LogWarning "svcValidation.ValidateChain", result.Message
    End If
    Set ValidateChain = result
End Function

' Changes one transaction amount inside a block (local Transactions table
' only) so Validate chain has something to catch. It picks the last
' transaction of the block (the coinbase if the block has no others).
Public Function TamperBlock(ByVal heightText As String) As clsResult
    Dim chain As Collection
    Dim blk As clsBlock
    Dim tx As clsTransaction
    Dim h As Long
    Dim oldAmount As Long
    Dim result As clsResult

    Set result = New clsResult
    Set chain = repoBlock.LoadChain()
    If Not modSheetUtils.TryParseLong(heightText, h) Then h = -1
    If h < 0 Or h > chain.Count - 1 Then
        result.Init False, modMessages.FormatText(modMessages.MSG_TAMPER_BAD_HEIGHT, heightText)
        Set TamperBlock = result
        Exit Function
    End If
    If h = 0 Then
        result.Init False, modMessages.MSG_TAMPER_GENESIS
        Set TamperBlock = result
        Exit Function
    End If

    Set blk = chain(h + 1)
    Set tx = blk.Transactions(blk.Transactions.Count)
    oldAmount = tx.Amount
    tx.Amount = oldAmount + 1
    repoTransaction.SaveChainTransactions chain
    clsLogger.LogWarning "svcValidation.TamperBlock", "Tampered with " & tx.TxID & " in block " & h
    result.Init True, modMessages.FormatText(modMessages.MSG_TAMPER_DONE, Left$(tx.TxID, 12), h, oldAmount, tx.Amount), 1
    Set TamperBlock = result
End Function
