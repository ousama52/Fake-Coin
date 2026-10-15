Attribute VB_Name = "svcConsensus"
Option Explicit

' Sync and the most-work rule. One Sync works on an in-memory copy of the
' chain and the Mempool, and writes the changed tables once at the end.
' Module-level variables below hold that working copy for the duration of one run.
Private mChain As Collection
Private mParams As Variant
Private mMempool As Collection
Private mSeen As Object
Private mIndex As Object
Private mFolderBlocks As Object
Private mLog As Collection
Private mChainChanged As Boolean
Private mMempoolChanged As Boolean
Private mCounts(0 To 5) As Long
Private mFiles As Long

Private Const IX_ADDED As Long = 0
Private Const IX_DUP As Long = 1
Private Const IX_REJECTED As Long = 2
Private Const IX_ORPHAN As Long = 3
Private Const IX_FORK As Long = 4
Private Const IX_REORG As Long = 5

' Reads new tx_ and block_ files, adds valid blocks and transactions, and
' switches to a chain with more cumulative work when one shows up.
Public Function Sync() As clsResult
    Dim names As Collection
    Dim i As Long
    Dim result As clsResult

    svcNetwork.EnsureFolder
    InitRun
    clsLogger.LogInfo "svcConsensus.Sync", "Sync started"
    Set names = svcNetwork.ListFileNames("block_")
    For i = 1 To names.Count
        ProcessBlockFile CStr(names(i))
    Next i
    Set names = svcNetwork.ListFileNames("tx_")
    For i = 1 To names.Count
        ProcessTxFile CStr(names(i))
    Next i
    CommitRun

    Set result = New clsResult
    result.Init True, modMessages.FormatText(modMessages.MSG_SYNC_DONE, mFiles, mCounts(IX_ADDED), mCounts(IX_DUP), _
        mCounts(IX_REJECTED), mCounts(IX_ORPHAN), mCounts(IX_FORK), mCounts(IX_REORG), mChain.Count - 1), mFiles
    clsLogger.LogInfo "svcConsensus.Sync", result.Message
    Set Sync = result
End Function

' Throws the local chain away (back to genesis, Mempool and SyncLog cleared)
' and rebuilds it from every file in the shared folder.
Public Function RebuildFromFolder() As clsResult
    Dim result As clsResult

    svcNetwork.EnsureFolder
    ResetLocalCore
    Set result = Sync()
    result.Message = modMessages.FormatText(modMessages.MSG_REBUILD_DONE, result.Message)
    Set RebuildFromFolder = result
End Function

' Back to the genesis block only. Wallets and key files are kept. SyncLog is
' cleared too, otherwise the next Sync would skip every file it already saw.
Public Function ResetLocalChain() As clsResult
    Dim result As clsResult

    ResetLocalCore
    Set result = New clsResult
    result.Init True, modMessages.MSG_RESET_LOCAL_DONE
    Set ResetLocalChain = result
End Function

Public Function ResetNetwork() As clsResult
    Dim n As Long
    Dim result As clsResult

    n = svcNetwork.DeleteAllFiles()
    clsLogger.LogWarning "svcConsensus.ResetNetwork", "Deleted " & n & " file(s)"
    Set result = New clsResult
    result.Init True, modMessages.FormatText(modMessages.MSG_RESET_NET_DONE, n), n
    Set ResetNetwork = result
End Function

Private Sub ResetLocalCore()
    Dim chain As Collection
    Dim none As Collection

    Set chain = New Collection
    chain.Add svcValidation.GenesisBlock()
    Set none = New Collection
    repoBlock.SaveChain chain
    repoMempool.SaveMempool none
    repoSyncLog.ClearSyncLog
    svcBalances.RebuildBalances chain
    clsLogger.LogWarning "svcConsensus.ResetLocalCore", "Local chain reset to genesis"
End Sub

' ---- Run set-up and commit -------------------------------------------------

Private Sub InitRun()
    Dim i As Long

    Set mChain = repoBlock.LoadChain()
    If mChain.Count = 0 Then mChain.Add svcValidation.GenesisBlock()
    mParams = repoConfig.LoadChainParams()
    Set mMempool = repoMempool.LoadMempool()
    Set mSeen = repoSyncLog.LoadSeenFiles()
    Set mFolderBlocks = Nothing
    Set mLog = New Collection
    mChainChanged = False
    mMempoolChanged = False
    mFiles = 0
    For i = 0 To 5
        mCounts(i) = 0
    Next i
    ReindexChain
End Sub

Private Sub ReindexChain()
    Dim i As Long

    Set mIndex = CreateObject("Scripting.Dictionary")
    For i = 1 To mChain.Count
        mIndex(mChain(i).Hash) = mChain(i).Height
    Next i
End Sub

' Writes only what changed: chain tables, Balances, Mempool, then the SyncLog rows.
Private Sub CommitRun()
    Dim rows() As Variant
    Dim data As Variant
    Dim i As Long

    If mChainChanged Then
        svcDifficulty.AssignCumulativeWork mChain
        repoBlock.SaveChain mChain
        svcBalances.RebuildBalances mChain
    End If
    If mMempoolChanged Then repoMempool.SaveMempool mMempool
    If mLog.Count > 0 Then
        ReDim rows(1 To mLog.Count, 1 To repoSyncLog.SLC_COUNT)
        For i = 1 To mLog.Count
            rows(i, repoSyncLog.SLC_FILE) = mLog(i)(0)
            rows(i, repoSyncLog.SLC_KIND) = mLog(i)(1)
            rows(i, repoSyncLog.SLC_ITEM) = mLog(i)(2)
            rows(i, repoSyncLog.SLC_RESULT) = mLog(i)(3)
            rows(i, repoSyncLog.SLC_REASON) = mLog(i)(4)
            rows(i, repoSyncLog.SLC_TIME) = mLog(i)(5)
        Next i
        data = rows
        repoSyncLog.AppendSyncLog data
    End If
End Sub

' Files already in SyncLog are skipped, except those whose last result was Orphan.
Private Function ShouldProcess(ByVal fileName As String) As Boolean
    If mSeen.Exists(fileName) Then
        ShouldProcess = (mSeen(fileName) = modConstants.SYNC_ORPHAN)
    Else
        ShouldProcess = True
    End If
End Function

Private Sub AddLog(ByVal fileName As String, ByVal kind As String, ByVal itemId As String, _
                   ByVal result As String, ByVal reason As String)
    mLog.Add Array(fileName, kind, itemId, result, reason, modTime.UnixToText(modTime.UnixNowUtc()))
    mSeen(fileName) = result
    Select Case result
        Case modConstants.SYNC_ADDED: mCounts(IX_ADDED) = mCounts(IX_ADDED) + 1
        Case modConstants.SYNC_DUPLICATE: mCounts(IX_DUP) = mCounts(IX_DUP) + 1
        Case modConstants.SYNC_REJECTED: mCounts(IX_REJECTED) = mCounts(IX_REJECTED) + 1
        Case modConstants.SYNC_ORPHAN: mCounts(IX_ORPHAN) = mCounts(IX_ORPHAN) + 1
        Case modConstants.SYNC_FORK_STORED: mCounts(IX_FORK) = mCounts(IX_FORK) + 1
        Case modConstants.SYNC_REORG: mCounts(IX_REORG) = mCounts(IX_REORG) + 1
    End Select
End Sub

' ---- Block files -----------------------------------------------------------

Private Sub ProcessBlockFile(ByVal fileName As String)
    Dim blk As clsBlock
    Dim reason As String

    If Not ShouldProcess(fileName) Then Exit Sub
    mFiles = mFiles + 1
    If Not svcNetwork.ReadBlockFile(fileName, blk, reason) Then
        AddLog fileName, modConstants.SYNC_KIND_BLOCK, "", modConstants.SYNC_REJECTED, reason
    ElseIf mIndex.Exists(blk.Hash) Then
        AddLog fileName, modConstants.SYNC_KIND_BLOCK, blk.Hash, modConstants.SYNC_DUPLICATE, ""
    Else
        HandleNewBlock fileName, blk
    End If
End Sub

Private Sub HandleNewBlock(ByVal fileName As String, ByVal blk As clsBlock)
    Dim tip As clsBlock

    Set tip = mChain(mChain.Count)
    If blk.PrevHash = tip.Hash Then
        ExtendTip fileName, blk
    Else
        HandleBranch fileName, blk
    End If
End Sub

' The block's parent is my tip: validate it against the chain and append it
' (a child of my tip always adds work, so it always wins).
Private Sub ExtendTip(ByVal fileName As String, ByVal blk As clsBlock)
    Dim st As Variant
    Dim reason As String

    st = svcBalances.StateFromChain(mChain)
    If Not svcValidation.CheckBlock(blk, mChain, mChain.Count, mParams, st, reason) Then
        AddLog fileName, modConstants.SYNC_KIND_BLOCK, blk.Hash, modConstants.SYNC_REJECTED, reason
        Exit Sub
    End If
    mChain.Add blk
    mIndex(blk.Hash) = blk.Height
    mChainChanged = True
    DropConfirmedFromMempool svcBalances.StateFromChain(mChain)
    AddLog fileName, modConstants.SYNC_KIND_BLOCK, blk.Hash, modConstants.SYNC_ADDED, ""
End Sub

' The block does not extend my tip. Follow PrevHash through the block files
' back to my chain (the fork point). Unknown parent = Orphan, retried next
' Sync. A valid branch with MORE cumulative work than my chain replaces it
' (a tie keeps mine); otherwise it is just remembered (the files stay in the folder).
Private Sub HandleBranch(ByVal fileName As String, ByVal blk As clsBlock)
    Dim branch As Collection
    Dim candidate As Collection
    Dim st As Variant
    Dim forkHeight As Long
    Dim badHeight As Long
    Dim reason As String

    Set branch = BuildBranch(blk, forkHeight, reason)
    If branch Is Nothing Then
        AddLog fileName, modConstants.SYNC_KIND_BLOCK, blk.Hash, modConstants.SYNC_ORPHAN, reason
        Exit Sub
    End If
    Set candidate = BuildCandidate(forkHeight, branch)
    If Not svcValidation.ReplayChain(candidate, mParams, st, badHeight, reason) Then
        AddLog fileName, modConstants.SYNC_KIND_BLOCK, blk.Hash, modConstants.SYNC_REJECTED, _
               "height " & badHeight & ": " & reason
    ElseIf svcDifficulty.ChainWork(candidate) > svcDifficulty.ChainWork(mChain) Then
        SwitchTo candidate, forkHeight, fileName, blk
    Else
        AddLog fileName, modConstants.SYNC_KIND_BLOCK, blk.Hash, modConstants.SYNC_FORK_STORED, ""
    End If
End Sub

' Blocks from the first block after the fork point up to blk. Nothing when a
' parent cannot be found in the folder.
Private Function BuildBranch(ByVal blk As clsBlock, ByRef forkHeight As Long, ByRef reason As String) As Collection
    Dim branch As Collection
    Dim cur As clsBlock
    Dim guard As Long

    LoadFolderBlocks
    Set branch = New Collection
    branch.Add blk
    Set cur = blk
    Do
        If mIndex.Exists(cur.PrevHash) Then
            forkHeight = mIndex(cur.PrevHash)
            Set BuildBranch = branch
            Exit Function
        End If
        guard = guard + 1
        If guard > mFolderBlocks.Count + 1 Then
            reason = modMessages.MSG_SYNC_LOOP
            Exit Function
        End If
        If Not mFolderBlocks.Exists(cur.PrevHash) Then
            reason = modMessages.FormatText(modMessages.MSG_SYNC_ORPHAN, Left$(cur.PrevHash, 12))
            Exit Function
        End If
        Set cur = mFolderBlocks(cur.PrevHash)
        branch.Add cur, , 1
    Loop
End Function

' My chain up to the fork point followed by the branch.
Private Function BuildCandidate(ByVal forkHeight As Long, ByVal branch As Collection) As Collection
    Dim cand As Collection
    Dim i As Long

    Set cand = New Collection
    For i = 1 To forkHeight + 1
        cand.Add mChain(i)
    Next i
    For i = 1 To branch.Count
        cand.Add branch(i)
    Next i
    Set BuildCandidate = cand
End Function

' Every block file with a readable, well-formed block, by hash (read once per Sync).
Private Sub LoadFolderBlocks()
    Dim names As Collection
    Dim blk As clsBlock
    Dim reason As String
    Dim i As Long

    If Not mFolderBlocks Is Nothing Then Exit Sub
    Set mFolderBlocks = CreateObject("Scripting.Dictionary")
    Set names = svcNetwork.ListFileNames("block_")
    For i = 1 To names.Count
        If svcNetwork.ReadBlockFile(CStr(names(i)), blk, reason) Then
            If Not mFolderBlocks.Exists(blk.Hash) Then mFolderBlocks.Add blk.Hash, blk
        End If
    Next i
End Sub

' Replaces my chain with the candidate: transactions of my dropped blocks that
' are not in the new chain go back to the Mempool (if still valid), confirmed
' Mempool entries are removed, and a Reorg line is logged.
Private Sub SwitchTo(ByVal candidate As Collection, ByVal forkHeight As Long, ByVal fileName As String, _
                     ByVal blk As clsBlock)
    Dim oldTip As clsBlock
    Dim oldWork As Variant
    Dim dropped As Collection
    Dim reason As String

    Set oldTip = mChain(mChain.Count)
    oldWork = svcDifficulty.ChainWork(mChain)
    Set dropped = DroppedTransactions(forkHeight)
    Set mChain = candidate
    ReindexChain
    DropConfirmedFromMempool svcBalances.StateFromChain(mChain)
    RequeueDropped dropped
    mChainChanged = True
    mMempoolChanged = True
    reason = modMessages.FormatText(modMessages.MSG_REORG, Left$(oldTip.Hash, 12), oldTip.Height, oldWork, _
                                    Left$(blk.Hash, 12), blk.Height, svcDifficulty.ChainWork(mChain), forkHeight)
    clsLogger.LogWarning "svcConsensus.SwitchTo", reason
    AddLog fileName, modConstants.SYNC_KIND_BLOCK, blk.Hash, modConstants.SYNC_REORG, reason
End Sub

' Non-coinbase transactions of my blocks above the fork point, in chain order.
Private Function DroppedTransactions(ByVal forkHeight As Long) As Collection
    Dim list As Collection
    Dim b As Long
    Dim t As Long
    Dim blk As clsBlock

    Set list = New Collection
    For b = forkHeight + 2 To mChain.Count
        Set blk = mChain(b)
        For t = 1 To blk.Transactions.Count
            If blk.Transactions(t).FromAddress <> modConstants.COINBASE_KEY Then list.Add blk.Transactions(t)
        Next t
    Next b
    Set DroppedTransactions = list
End Function

' Puts dropped transactions that are not in the new chain back into the
' Mempool when they pass every admission rule against the new chain.
Private Sub RequeueDropped(ByVal dropped As Collection)
    Dim st As Variant
    Dim tx As clsTransaction
    Dim reason As String
    Dim i As Long

    st = svcBalances.StateFromChain(mChain)
    For i = 1 To dropped.Count
        Set tx = dropped(i)
        tx.Source = modConstants.SRC_NETWORK
        If svcTransaction.AdmitToMempool(tx, mMempool, st, mParams, mChain.Count - 1, reason) = modConstants.ADM_OK Then
            mMempoolChanged = True
        End If
    Next i
End Sub

Private Sub DropConfirmedFromMempool(ByRef st As Variant)
    Dim rest As Collection
    Dim i As Long

    Set rest = New Collection
    For i = 1 To mMempool.Count
        If Not st(2).Exists(mMempool(i).TxID) Then rest.Add mMempool(i)
    Next i
    If rest.Count <> mMempool.Count Then
        Set mMempool = rest
        mMempoolChanged = True
    End If
End Sub

' ---- Transaction files -----------------------------------------------------

' A transaction from the network joins the Mempool when it passes every
' admission rule (signature included) on top of the confirmed chain and this
' sender's pending transactions. A nonce that is ahead of the sender's pending
' count is an Orphan: it is retried on the next Sync.
Private Sub ProcessTxFile(ByVal fileName As String)
    Dim tx As clsTransaction
    Dim st As Variant
    Dim reason As String
    Dim code As Long

    If Not ShouldProcess(fileName) Then Exit Sub
    mFiles = mFiles + 1
    If Not svcNetwork.ReadTxFile(fileName, tx, reason) Then
        AddLog fileName, modConstants.SYNC_KIND_TX, "", modConstants.SYNC_REJECTED, reason
        Exit Sub
    End If
    st = svcBalances.StateFromChain(mChain)
    tx.Source = modConstants.SRC_NETWORK
    code = svcTransaction.AdmitToMempool(tx, mMempool, st, mParams, mChain.Count - 1, reason)
    Select Case code
        Case modConstants.ADM_OK
            mMempoolChanged = True
            AddLog fileName, modConstants.SYNC_KIND_TX, tx.TxID, modConstants.SYNC_ADDED, ""
        Case modConstants.ADM_DUP
            AddLog fileName, modConstants.SYNC_KIND_TX, tx.TxID, modConstants.SYNC_DUPLICATE, ""
        Case modConstants.ADM_WAIT
            AddLog fileName, modConstants.SYNC_KIND_TX, tx.TxID, modConstants.SYNC_ORPHAN, reason
        Case Else
            AddLog fileName, modConstants.SYNC_KIND_TX, tx.TxID, modConstants.SYNC_REJECTED, reason
    End Select
End Sub
