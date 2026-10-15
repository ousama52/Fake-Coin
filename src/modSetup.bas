Attribute VB_Name = "modSetup"
Option Explicit

' FOR LEARNING ONLY.
' Creates whatever is missing (sheets, tables, columns, config keys, genesis
' block, balances, Control sheet and buttons). Nothing that already exists is
' cleared or overwritten, so it is safe to run again.

Public Function EnsureAll() As clsResult
    Dim created As Long
    Dim result As clsResult

    ' Tables first: the logger writes to tblLog
    created = EnsureTables()
    clsLogger.LogInfo "modSetup.EnsureAll", "Setup started, hash engine: " & svcHash.HashEngine()
    created = created + EnsureConfigDefaults()
    created = created + EnsureGenesis()
    created = created + EnsureBalances()
    EnsureWalletsNote
    EnsureControlSheet
    clsLogger.LogInfo "modSetup.EnsureAll", "Setup finished, items created: " & created

    Set result = New clsResult
    result.Init True, modMessages.FormatText(modMessages.MSG_SETUP_DONE, created), created
    Set EnsureAll = result
End Function

' Column spec: Header:Type[:EnumList]
' Types: T text (Text format), N whole number, E enum (list validation)
Private Function TableSpecs() As Variant
    TableSpecs = Array( _
        modConstants.TBL_LOG & "=Timestamp:T|Level:E:Info,Warning,Error|Procedure:T|Message:T", _
        modConstants.TBL_CONFIG & "=Key:T|Value:T", _
        modConstants.TBL_WALLETS & "=Name:T|Address:T|PublicKey:T|KeyFile:T|CreatedAt:N", _
        modConstants.TBL_MEMPOOL & "=TxID:T|FromAddress:T|FromPubKey:T|ToAddress:T|Amount:N|Fee:N|Nonce:N|Timestamp:N|Signature:T|Source:E:Local,Network", _
        modConstants.TBL_BLOCKS & "=Height:N|Timestamp:N|PrevHash:T|MerkleRoot:T|Nonce:N|DifficultyBits:N|Miner:T|TxCount:N|Hash:T|CumulativeWork:N|Source:E:Genesis,Mined,Network", _
        modConstants.TBL_TRANSACTIONS & "=TxID:T|BlockHeight:N|Position:N|FromAddress:T|FromPubKey:T|ToAddress:T|Amount:N|Fee:N|Nonce:N|Timestamp:N|Signature:T", _
        modConstants.TBL_BALANCES & "=Address:T|Name:T|Balance:N|Spendable:N|NextNonce:N", _
        modConstants.TBL_SYNCLOG & "=FileName:T|Kind:E:Block,Tx|ItemID:T|Result:E:Added,Duplicate,Rejected,Orphan,ForkStored,Reorg|Reason:T|ProcessedAt:T")
End Function

Private Function EnsureTables() As Long
    Dim specs As Variant
    Dim i As Long

    specs = TableSpecs()
    For i = LBound(specs) To UBound(specs)
        EnsureTables = EnsureTables + EnsureOneTable(CStr(specs(i)))
    Next i
End Function

' Returns the number of items created (the table and any added columns).
Private Function EnsureOneTable(ByVal spec As String) As Long
    Dim tblName As String
    Dim colSpecs As Variant
    Dim headers As Variant
    Dim ws As Worksheet
    Dim lo As ListObject
    Dim existed As Boolean
    Dim added As Long
    Dim i As Long

    tblName = Left$(spec, InStr(spec, "=") - 1)
    colSpecs = Split(Mid$(spec, InStr(spec, "=") + 1), "|")
    headers = HeadersOf(colSpecs)

    existed = modSheetUtils.TableExists(tblName)
    Set ws = modSheetUtils.GetOrCreateSheet(Mid$(tblName, 4))
    Set lo = modSheetUtils.EnsureTable(ws, tblName, headers)
    added = modSheetUtils.AddMissingColumns(lo, headers)
    For i = LBound(colSpecs) To UBound(colSpecs)
        ApplyColumnRules lo, CStr(colSpecs(i))
    Next i
    If Not existed Then EnsureOneTable = 1
    EnsureOneTable = EnsureOneTable + added
End Function

Private Function HeadersOf(ByVal colSpecs As Variant) As Variant
    Dim result() As Variant
    Dim i As Long

    ReDim result(0 To UBound(colSpecs))
    For i = 0 To UBound(colSpecs)
        result(i) = Split(CStr(colSpecs(i)), ":")(0)
    Next i
    HeadersOf = result
End Function

Private Sub ApplyColumnRules(ByVal lo As ListObject, ByVal colSpec As String)
    Dim parts As Variant
    Dim idx As Long

    parts = Split(colSpec, ":")
    idx = modSheetUtils.ColIndex(lo, CStr(parts(0)))
    If idx = 0 Then Exit Sub
    ' Hash, key, address and signature columns stay Text so Excel never changes them
    Select Case CStr(parts(1))
        Case "T", "E": modSheetUtils.SetColumnFormat lo, idx, "@"
        Case "N": modSheetUtils.SetColumnFormat lo, idx, "0"
    End Select
    If CStr(parts(1)) = "E" Then
        modSheetUtils.AddListValidation modSheetUtils.ColumnBelowHeader(lo, idx), CStr(parts(2))
    End If
End Sub

' ---- Config, genesis, balances ---------------------------------------------

Private Function ConfigDefaults() As Variant
    ConfigDefaults = modSheetUtils.JaggedToTable(Array( _
        Array(modConstants.CFG_COIN_NAME, "FakeCoin"), _
        Array(modConstants.CFG_NODE_ID, "NodeA"), _
        Array(modConstants.CFG_SHARED_FOLDER, "C:\FakeCoin\network"), _
        Array(modConstants.CFG_WALLET_FOLDER, "C:\FakeCoin\wallets"), _
        Array(modConstants.CFG_BLOCK_REWARD, "50"), _
        Array(modConstants.CFG_HALVING, "10"), _
        Array(modConstants.CFG_MAX_TX, "5"), _
        Array(modConstants.CFG_MIN_FEE, "1"), _
        Array(modConstants.CFG_INIT_BITS, "12"), _
        Array(modConstants.CFG_TARGET_SECONDS, "30"), _
        Array(modConstants.CFG_DIFF_INTERVAL, "10"), _
        Array(modConstants.CFG_MIN_BITS, "4"), _
        Array(modConstants.CFG_MAX_BITS, "32"), _
        Array(modConstants.CFG_MATURITY, "3"), _
        Array(modConstants.CFG_MAX_FUTURE, "600"), _
        Array(modConstants.CFG_MEMPOOL_MAX, "50")))
End Function

' Adds only the keys that are not in the table yet; existing values stay.
Private Function EnsureConfigDefaults() As Long
    Dim lo As ListObject
    Dim current As Variant
    Dim defaults As Variant
    Dim missing() As Variant
    Dim isMissing() As Boolean
    Dim i As Long
    Dim n As Long

    Set lo = modSheetUtils.GetTable(modConstants.TBL_CONFIG)
    current = modSheetUtils.ReadTable(lo)
    defaults = ConfigDefaults()
    ReDim isMissing(1 To UBound(defaults, 1))
    For i = 1 To UBound(defaults, 1)
        isMissing(i) = Not KeyExists(current, CStr(defaults(i, 1)))
        If isMissing(i) Then n = n + 1
    Next i
    If n = 0 Then Exit Function

    ReDim missing(1 To n, 1 To 2)
    n = 0
    For i = 1 To UBound(defaults, 1)
        If isMissing(i) Then
            n = n + 1
            missing(n, 1) = defaults(i, 1)
            missing(n, 2) = defaults(i, 2)
        End If
    Next i
    modSheetUtils.AppendRows lo, missing
    clsLogger.LogInfo "modSetup.EnsureConfigDefaults", "Added " & n & " config key(s)"
    EnsureConfigDefaults = n
End Function

Private Function KeyExists(ByRef data As Variant, ByVal key As String) As Boolean
    Dim i As Long

    If modSheetUtils.RowCount(data) = 0 Then Exit Function
    For i = LBound(data, 1) To UBound(data, 1)
        If StrComp(modSheetUtils.CellText(data(i, 1)), key, vbTextCompare) = 0 Then
            KeyExists = True
            Exit Function
        End If
    Next i
End Function

' Writes the fixed genesis block when Blocks is empty, and always checks that
' the hash function reproduces the genesis hash. If it does not, the hash
' function is wrong and nothing else can work.
Private Function EnsureGenesis() As Long
    Dim chain As Collection
    Dim computed As String

    computed = svcValidation.BlockHashOf(svcValidation.GenesisBlock())
    If computed <> modConstants.GENESIS_HASH Then
        Err.Raise modConstants.ERR_APP + 50, "modSetup.EnsureGenesis", _
                  modMessages.FormatText(modMessages.MSG_GENESIS_BAD, computed, modConstants.GENESIS_HASH)
    End If
    Set chain = repoBlock.LoadChain()
    If chain.Count > 0 Then Exit Function
    chain.Add svcValidation.GenesisBlock()
    svcDifficulty.AssignCumulativeWork chain
    repoBlock.SaveChain chain
    clsLogger.LogInfo "modSetup.EnsureGenesis", "Genesis block written"
    EnsureGenesis = 1
End Function

' Balances is empty until there is a wallet or a transaction; always rebuilt, never edited.
Private Function EnsureBalances() As Long
    If Not modSheetUtils.TableIsEmpty(modSheetUtils.GetTable(modConstants.TBL_BALANCES)) Then Exit Function
    svcBalances.RebuildBalances repoBlock.LoadChain()
    If Not modSheetUtils.TableIsEmpty(modSheetUtils.GetTable(modConstants.TBL_BALANCES)) Then EnsureBalances = 1
End Function

Private Sub EnsureWalletsNote()
    Dim ws As Worksheet

    Set ws = modSheetUtils.GetSheet(modConstants.SHEET_WALLETS)
    ' Column H: two empty columns after the table so it never grows into the note
    ws.Range("H1").Value = modMessages.LBL_FAKE
    ws.Range("H1").Font.Bold = True
    ws.Range("H1").Font.Color = RGB(192, 0, 0)
End Sub

' ---- Control sheet ---------------------------------------------------------

Private Sub EnsureControlSheet()
    Dim ws As Worksheet
    Dim isNew As Boolean

    isNew = Not modSheetUtils.SheetExists(modConstants.SHEET_CONTROL)
    Set ws = modSheetUtils.GetOrCreateSheet(modConstants.SHEET_CONTROL)
    If isNew Then ws.Move Before:=ThisWorkbook.Worksheets(1)
    WriteControlLabels ws
    RemoveStaleButtons ws
    AddControlButtons ws
End Sub

' Labels are written every time (they hold no user input). The input cells
' C4:C13 are never touched.
Private Sub WriteControlLabels(ByVal ws As Worksheet)
    ws.Range("B1").Value = modMessages.LBL_CTL_TITLE
    ws.Range("B1").Font.Bold = True
    ws.Range("B1").Font.Size = 16
    ws.Range("B2").Value = modMessages.LBL_FAKE
    ws.Range("B2").Font.Bold = True
    ws.Range("B2").Font.Color = RGB(192, 0, 0)
    ws.Range("B4").Value = modMessages.LBL_CTL_FROM
    ws.Range("B5").Value = modMessages.LBL_CTL_TO
    ws.Range("B6").Value = modMessages.LBL_CTL_AMOUNT
    ws.Range("B7").Value = modMessages.LBL_CTL_FEE
    ws.Range("B8").Value = modMessages.LBL_CTL_MINER
    ws.Range("B9").Value = modMessages.LBL_CTL_HEIGHT
    ws.Range("B10").Value = modMessages.LBL_CTL_WALLET
    ws.Range("B11").Value = modMessages.LBL_CTL_STATUS
    ws.Range("B12").Value = modMessages.LBL_CTL_TXID
    ws.Range("B13").Value = modMessages.LBL_CTL_NONCE
    ' Names, addresses and TxIDs are Text so they stay exactly as typed
    ws.Range("C4:C5,C8,C10,C12").NumberFormat = "@"
    ws.Range("C4:C10,C12:C13").Interior.Color = RGB(255, 255, 204)
    ws.Columns("A").ColumnWidth = 3
    ws.Columns("B").ColumnWidth = 38
    ws.Columns("C").ColumnWidth = 44
    ws.Columns("D").ColumnWidth = 3
End Sub

' A button from an earlier build of this workbook that no longer has a macro.
Private Sub RemoveStaleButtons(ByVal ws As Worksheet)
    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes("btnUpgrade")
    If Not shp Is Nothing Then shp.Delete
End Sub

Private Sub AddControlButtons(ByVal ws As Worksheet)
    Dim specs As Variant
    Dim i As Long
    Dim parts As Variant
    Dim btn As Object

    specs = Array( _
        "btnSetup|" & modMessages.BTN_SETUP & "|modMain.RunSetup", _
        "btnCreateWallet|" & modMessages.BTN_CREATE_WALLET & "|modMain.RunCreateWallet", _
        "btnSend|" & modMessages.BTN_SEND & "|modMain.RunSendCoins", _
        "btnMine|" & modMessages.BTN_MINE & "|modMain.RunMineBlock", _
        "btnSync|" & modMessages.BTN_SYNC & "|modMain.RunSync", _
        "btnValidate|" & modMessages.BTN_VALIDATE & "|modMain.RunValidateChain", _
        "btnRebuild|" & modMessages.BTN_REBUILD & "|modMain.RunRebuildFromFolder", _
        "btnTamper|" & modMessages.BTN_TAMPER & "|modMain.RunTamperTest", _
        "btnTxStatus|" & modMessages.BTN_TXSTATUS & "|modMain.RunTxStatus", _
        "btnForge|" & modMessages.BTN_FORGE & "|modMain.RunForgeTest", _
        "btnResetLocal|" & modMessages.BTN_RESET_LOCAL & "|modMain.RunResetLocal", _
        "btnResetNet|" & modMessages.BTN_RESET_NET & "|modMain.RunResetNetwork")
    For i = LBound(specs) To UBound(specs)
        parts = Split(CStr(specs(i)), "|")
        If Not ButtonExists(ws, CStr(parts(0))) Then
            Set btn = ws.Buttons.Add(ws.Range("E4").Left, ws.Range("E4").Top + i * 28, 170, 24)
            btn.Name = CStr(parts(0))
            btn.Caption = CStr(parts(1))
            btn.OnAction = CStr(parts(2))
        End If
    Next i
End Sub

Private Function ButtonExists(ByVal ws As Worksheet, ByVal btnName As String) As Boolean
    Dim shp As Shape

    On Error Resume Next
    Set shp = ws.Shapes(btnName)
    ButtonExists = Not shp Is Nothing
End Function
