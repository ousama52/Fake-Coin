Attribute VB_Name = "svcMining"
Option Explicit

' Mines one block from the Mempool. Esc stops the search; nothing is written
' unless a valid block was found. The block is built and checked in memory,
' then the tables and the block file are written once at the end.
Public Function MineBlock(ByVal minerInput As String) As clsResult
    Dim result As clsResult
    Dim n As Long
    Dim src As String
    Dim desc As String

    On Error GoTo Fail
    Application.EnableCancelKey = xlErrorHandler
    Set result = MineInternal(minerInput)
    Application.EnableCancelKey = xlInterrupt
    Set MineBlock = result
    Exit Function
Fail:
    n = Err.Number
    src = Err.Source
    desc = Err.Description
    Application.EnableCancelKey = xlInterrupt
    If n = 18 Then
        Set result = New clsResult
        result.Init False, modMessages.MSG_MINING_STOPPED
        Set MineBlock = result
    Else
        Err.Raise n, src, desc
    End If
End Function

Private Function MineInternal(ByVal minerInput As String) As clsResult
    Dim params As Variant
    Dim chain As Collection
    Dim mem As Collection
    Dim picked As Collection
    Dim tip As clsBlock
    Dim blk As clsBlock
    Dim state As Variant
    Dim minerAddr As String
    Dim reason As String
    Dim result As clsResult

    minerAddr = svcWallet.ResolveAddress(minerInput)
    Set result = New clsResult
    If Len(minerAddr) = 0 Then
        result.Init False, modMessages.FormatText(modMessages.MSG_MINER_UNKNOWN, minerInput)
        Set MineInternal = result
        Exit Function
    End If

    params = repoConfig.LoadChainParams()
    Set chain = repoBlock.LoadChain()
    Set tip = chain(chain.Count)
    state = svcBalances.StateFromChain(chain)
    Set mem = repoMempool.LoadMempool()
    Set picked = SelectTransactions(mem, params, state, tip.Height + 1)
    Set blk = BuildBlock(tip, chain, minerAddr, picked, params)
    clsLogger.LogInfo "svcMining.MineBlock", "Mining height " & blk.Height & " at " & blk.DifficultyBits & " bits with " & picked.Count & " transaction(s)"
    FindNonce blk

    ' Same rules every other node will apply, against a fresh replay
    state = svcBalances.StateFromChain(chain)
    If Not svcValidation.CheckBlock(blk, chain, chain.Count, params, state, reason) Then
        Err.Raise modConstants.ERR_APP + 20, "svcMining.MineBlock", "Mined block failed its own check: " & reason
    End If

    svcNetwork.EnsureFolder
    chain.Add blk
    svcDifficulty.AssignCumulativeWork chain
    repoBlock.SaveChain chain
    repoMempool.SaveMempool WithoutMined(mem, blk)
    svcBalances.RebuildBalances chain
    svcNetwork.WriteBlockFile blk
    clsLogger.LogInfo "svcMining.MineBlock", "Mined block " & blk.Height & " hash " & blk.Hash

    result.Init True, modMessages.FormatText(modMessages.MSG_MINED, blk.Height, Left$(blk.Hash, 16), blk.DifficultyBits, _
        blk.TxCount, blk.Nonce, blk.Transactions(1).Amount), blk.Height
    Set MineInternal = result
End Function

' Mempool transactions that pass the rules in order (signature, nonce,
' spendable balance at the new height), highest fee first, up to
' MaxTxPerBlock. After each pick the scan restarts, so a sender's nonce 1 can
' follow its nonce 0 even when nonce 1 pays the higher fee.
Private Function SelectTransactions(ByVal mem As Collection, ByRef params As Variant, _
                                    ByRef state As Variant, ByVal height As Long) As Collection
    Dim remaining As Collection
    Dim picked As Collection
    Dim i As Long
    Dim found As Boolean
    Dim reason As String
    Dim tx As clsTransaction

    Set remaining = SortByFeeDescending(mem)
    Set picked = New Collection
    Do
        found = False
        For i = 1 To remaining.Count
            Set tx = remaining(i)
            If svcTransaction.CheckTransaction(tx, params, state, height, reason) Then
                svcBalances.ApplyTx tx, state
                picked.Add tx
                remaining.Remove i
                found = True
                Exit For
            End If
        Next i
    Loop While found And picked.Count < params(modConstants.PARAM_MAXTX)
    Set SelectTransactions = picked
End Function

' Stable insertion sort: equal fees keep their Mempool order.
Private Function SortByFeeDescending(ByVal mem As Collection) As Collection
    Dim sorted As Collection
    Dim i As Long
    Dim j As Long
    Dim tx As clsTransaction
    Dim placed As Boolean

    Set sorted = New Collection
    For i = 1 To mem.Count
        Set tx = mem(i)
        placed = False
        For j = 1 To sorted.Count
            If sorted(j).Fee < tx.Fee Then
                sorted.Add tx, , j
                placed = True
                Exit For
            End If
        Next j
        If Not placed Then sorted.Add tx
    Next i
    Set SortByFeeDescending = sorted
End Function

' Coinbase first (reward + fees of the picked transactions, paid to any
' address), then the picked transactions, with the Merkle root. The timestamp
' is the current UTC time, but always later than the parent's. The required
' difficulty comes from the chain. The nonce and hash are filled by FindNonce.
Private Function BuildBlock(ByVal tip As clsBlock, ByVal chain As Collection, ByVal minerAddr As String, _
                            ByVal picked As Collection, ByRef params As Variant) As clsBlock
    Dim blk As clsBlock
    Dim cb As clsTransaction
    Dim ids As Collection
    Dim fees As Long
    Dim i As Long

    For i = 1 To picked.Count
        fees = fees + picked(i).Fee
    Next i
    Set blk = New clsBlock
    blk.Height = tip.Height + 1
    blk.Timestamp = modTime.UnixNowUtc()
    If blk.Timestamp <= tip.Timestamp Then blk.Timestamp = tip.Timestamp + 1
    blk.PrevHash = tip.Hash
    blk.DifficultyBits = svcDifficulty.RequiredBits(blk.Height, chain, params)
    blk.Miner = minerAddr
    blk.Source = modConstants.SRC_MINED

    Set cb = svcTransaction.NewCoinbase(blk.Height, minerAddr, _
                svcValidation.BlockReward(blk.Height, params) + fees, blk.Timestamp)
    blk.Transactions.Add cb
    For i = 1 To picked.Count
        picked(i).BlockHeight = blk.Height
        picked(i).Position = i
        blk.Transactions.Add picked(i)
    Next i
    blk.TxCount = blk.Transactions.Count

    Set ids = New Collection
    For i = 1 To blk.Transactions.Count
        ids.Add blk.Transactions(i).TxID
    Next i
    blk.MerkleRoot = svcMerkle.MerkleRoot(ids)
    Set BuildBlock = blk
End Function

' Proof of work: count the nonce up from 0 until the hash has at least
' DifficultyBits leading zero bits. DoEvents every 5000 tries keeps Excel responsive.
Private Sub FindNonce(ByVal blk As clsBlock)
    Dim nonce As Long
    Dim h As String

    Do
        h = svcValidation.ComputeBlockHash(blk.Height, blk.Timestamp, blk.PrevHash, blk.MerkleRoot, _
                                           nonce, blk.DifficultyBits, blk.Miner)
        If svcHash.LeadingZeroBits(h) >= blk.DifficultyBits Then Exit Do
        nonce = nonce + 1
        If nonce Mod 5000 = 0 Then DoEvents
    Loop
    blk.Nonce = nonce
    blk.Hash = h
End Sub

' The Mempool without the transactions that are now in the block.
Private Function WithoutMined(ByVal mem As Collection, ByVal blk As clsBlock) As Collection
    Dim mined As Object
    Dim rest As Collection
    Dim i As Long

    Set mined = CreateObject("Scripting.Dictionary")
    For i = 1 To blk.Transactions.Count
        mined(blk.Transactions(i).TxID) = True
    Next i
    Set rest = New Collection
    For i = 1 To mem.Count
        If Not mined.Exists(mem(i).TxID) Then rest.Add mem(i)
    Next i
    Set WithoutMined = rest
End Function
