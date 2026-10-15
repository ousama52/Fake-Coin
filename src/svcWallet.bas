Attribute VB_Name = "svcWallet"
Option Explicit

' ============================================================================
' FOR LEARNING ONLY.
' A wallet is a name, an address, a public key and a KEY FILE. The private key
' lives only in the key file <WalletFolder>\<NodeID>\<name>.key and is never
' written into the workbook or the shared folder. The key file is NOT encrypted.
' Each node lists only its own wallets; mining is separate from wallets.
' ============================================================================

Public Function CreateWallet(ByVal walletName As String) As clsResult
    Dim path As String
    Dim privateKey As String
    Dim publicKey As String
    Dim address As String
    Dim result As clsResult

    walletName = Trim$(walletName)
    If Len(walletName) = 0 Then
        Set CreateWallet = Fail(modMessages.MSG_WALLET_NAME_EMPTY)
        Exit Function
    End If
    If walletName Like "*[!A-Za-z0-9_]*" Then
        Set CreateWallet = Fail(modMessages.MSG_WALLET_BADNAME)
        Exit Function
    End If
    If repoWallet.FindWalletByName(repoWallet.LoadWallets(), walletName) > 0 Then
        Set CreateWallet = Fail(modMessages.FormatText(modMessages.MSG_WALLET_EXISTS, walletName))
        Exit Function
    End If
    path = KeyFilePath(walletName)
    If FileExists(path) Then
        Set CreateWallet = Fail(modMessages.FormatText(modMessages.MSG_WALLET_KEYFILE_EXISTS, path))
        Exit Function
    End If

    svcCrypto.GenerateKeyPair privateKey, publicKey
    address = svcCrypto.AddressFromPublicKey(publicKey)
    WriteKeyFile path, privateKey
    repoWallet.AppendWallet walletName, address, publicKey, path, modTime.UnixNowUtc()
    svcBalances.RebuildBalances repoBlock.LoadChain()
    clsLogger.LogInfo "svcWallet.CreateWallet", "Wallet " & walletName & " created, address " & address

    Set result = New clsResult
    result.Init True, modMessages.FormatText(modMessages.MSG_WALLET_CREATED, walletName, address, path), 1
    Set CreateWallet = result
End Function

' FC followed by 40 lowercase hex characters.
Public Function IsAddress(ByVal s As String) As Boolean
    If Len(s) <> modConstants.ADDRESS_LENGTH Then Exit Function
    If Left$(s, Len(modConstants.ADDRESS_PREFIX)) <> modConstants.ADDRESS_PREFIX Then Exit Function
    IsAddress = Not (Mid$(s, Len(modConstants.ADDRESS_PREFIX) + 1) Like "*[!0-9a-f]*")
End Function

' A wallet name of this node or any address becomes an address. "" when unknown.
Public Function ResolveAddress(ByVal nameOrAddress As String) As String
    Dim wallets As Variant
    Dim r As Long

    nameOrAddress = Trim$(nameOrAddress)
    If Len(nameOrAddress) = 0 Then Exit Function
    wallets = repoWallet.LoadWallets()
    r = repoWallet.FindWalletByName(wallets, nameOrAddress)
    If r > 0 Then
        ResolveAddress = modSheetUtils.CellText(wallets(r, repoWallet.WLC_ADDRESS))
    ElseIf IsAddress(nameOrAddress) Then
        ResolveAddress = nameOrAddress
    End If
End Function

' Private key text of one of this node's wallets, read from its key file.
' Returns "" and sets reason when the wallet or its key file is missing.
Public Function LoadPrivateKey(ByVal walletName As String, ByRef address As String, _
                               ByRef publicKey As String, ByRef reason As String) As String
    Dim wallets As Variant
    Dim r As Long
    Dim key As String

    wallets = repoWallet.LoadWallets()
    r = repoWallet.FindWalletByName(wallets, walletName)
    If r = 0 Then
        reason = modMessages.FormatText(modMessages.MSG_TX_R_FROM, walletName)
        Exit Function
    End If
    address = modSheetUtils.CellText(wallets(r, repoWallet.WLC_ADDRESS))
    publicKey = modSheetUtils.CellText(wallets(r, repoWallet.WLC_PUBKEY))
    key = ReadKeyFile(modSheetUtils.CellText(wallets(r, repoWallet.WLC_KEYFILE)))
    If Len(key) = 0 Then
        reason = modMessages.MSG_WALLET_NOKEY
        Exit Function
    End If
    LoadPrivateKey = key
End Function

' <WalletFolder>\<NodeID>\<name>.key
Public Function KeyFilePath(ByVal walletName As String) As String
    Dim folder As String

    folder = repoConfig.ConfigText(modConstants.CFG_WALLET_FOLDER, "C:\FakeCoin\wallets")
    Do While Len(folder) > 3 And Right$(folder, 1) = "\"
        folder = Left$(folder, Len(folder) - 1)
    Loop
    KeyFilePath = folder & "\" & repoConfig.ConfigText(modConstants.CFG_NODE_ID, "NodeA") & "\" & walletName & ".key"
End Function

' ---- Key files -------------------------------------------------------------

Private Function FileExists(ByVal path As String) As Boolean
    FileExists = CreateObject("Scripting.FileSystemObject").FileExists(path)
End Function

' Creates the folders, then writes the key file (never over an existing one).
Private Sub WriteKeyFile(ByVal path As String, ByVal privateKey As String)
    Dim f As Object
    Dim ts As Object

    Set f = CreateObject("Scripting.FileSystemObject")
    EnsureFolders f, f.GetParentFolderName(path)
    Set ts = f.CreateTextFile(path, False, False)
    ts.Write modConstants.KEY_FILE_HEADER & vbCrLf & privateKey & vbCrLf
    ts.Close
End Sub

Private Sub EnsureFolders(ByVal f As Object, ByVal folder As String)
    Dim parent As String

    If Len(folder) = 0 Or f.FolderExists(folder) Then Exit Sub
    parent = f.GetParentFolderName(folder)
    If Len(parent) > 0 Then EnsureFolders f, parent
    f.CreateFolder folder
End Sub

' The private key text in a key file, "" when the file is missing or not a key file.
Private Function ReadKeyFile(ByVal path As String) As String
    Dim f As Object
    Dim ts As Object
    Dim text As String
    Dim parts() As String

    Set f = CreateObject("Scripting.FileSystemObject")
    If Len(path) = 0 Then Exit Function
    If Not f.FileExists(path) Then Exit Function
    Set ts = f.OpenTextFile(path, 1, False, 0)
    If Not ts.AtEndOfStream Then text = ts.ReadAll
    ts.Close
    parts = Split(Replace(text, vbCrLf, vbLf), vbLf)
    If UBound(parts) < 1 Then Exit Function
    If parts(0) <> modConstants.KEY_FILE_HEADER Then Exit Function
    ReadKeyFile = Trim$(parts(1))
End Function

Private Function Fail(ByVal msg As String) As clsResult
    Dim r As clsResult

    Set r = New clsResult
    r.Init False, msg
    Set Fail = r
End Function
