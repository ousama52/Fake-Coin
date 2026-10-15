Attribute VB_Name = "svcNetwork"
Option Explicit

' The only module that touches the shared folder (the fake network).
' Files are plain ASCII, CRLF, and are never overwritten. A file is written
' under a temporary name and then renamed, so another node never sees a
' half-written block_ or tx_ file. A file with a wrong header is rejected.

Private Function Fso() As Object
    Set Fso = CreateObject("Scripting.FileSystemObject")
End Function

Public Function SharedFolder() As String
    Dim p As String

    p = repoConfig.ConfigText(modConstants.CFG_SHARED_FOLDER, "")
    Do While Len(p) > 3 And Right$(p, 1) = "\"
        p = Left$(p, Len(p) - 1)
    Loop
    SharedFolder = p
End Function

Public Sub EnsureFolder()
    Dim p As String

    p = SharedFolder()
    If Len(p) = 0 Or Not Fso().FolderExists(p) Then
        Err.Raise modConstants.ERR_APP + 40, "svcNetwork.EnsureFolder", _
                  modMessages.FormatText(modMessages.MSG_FOLDER_MISSING, p)
    End If
End Sub

' ---- Writing ---------------------------------------------------------------

' False when a file with that name already exists (it is left untouched).
Public Function WriteTxFile(ByVal tx As clsTransaction) As Boolean
    Dim text As String

    text = modConstants.FILE_TX_HEADER & vbCrLf & TxLine(tx) & vbCrLf
    WriteTxFile = WriteNewFile("tx_" & Left$(tx.TxID, 16) & ".txt", text)
End Function

Public Function WriteBlockFile(ByVal blk As clsBlock) As Boolean
    Dim text As String
    Dim i As Long

    text = modConstants.FILE_BLOCK_HEADER & vbCrLf & _
           "Height=" & blk.Height & vbCrLf & _
           "Timestamp=" & blk.Timestamp & vbCrLf & _
           "PrevHash=" & blk.PrevHash & vbCrLf & _
           "MerkleRoot=" & blk.MerkleRoot & vbCrLf & _
           "Nonce=" & blk.Nonce & vbCrLf & _
           "DifficultyBits=" & blk.DifficultyBits & vbCrLf & _
           "Miner=" & blk.Miner & vbCrLf & _
           "Hash=" & blk.Hash & vbCrLf & _
           "TxCount=" & blk.TxCount & vbCrLf
    For i = 1 To blk.Transactions.Count
        text = text & TxLine(blk.Transactions(i)) & vbCrLf
    Next i
    WriteBlockFile = WriteNewFile("block_" & Format$(blk.Height, "000000") & "_" & Left$(blk.Hash, 12) & ".txt", text)
End Function

' TX|TxID|FromAddress|FromPubKey|ToAddress|Amount|Fee|Nonce|Timestamp|Signature
' (a coinbase leaves FromPubKey and Signature empty but keeps the separators)
Private Function TxLine(ByVal tx As clsTransaction) As String
    TxLine = "TX|" & tx.TxID & "|" & tx.FromAddress & "|" & tx.FromPubKey & "|" & tx.ToAddress & "|" & _
             tx.Amount & "|" & tx.Fee & "|" & tx.Nonce & "|" & tx.Timestamp & "|" & tx.Signature
End Function

Private Function WriteNewFile(ByVal fileName As String, ByVal text As String) As Boolean
    Dim f As Object
    Dim finalPath As String
    Dim tempPath As String
    Dim ts As Object

    Set f = Fso()
    finalPath = SharedFolder() & "\" & fileName
    If f.FileExists(finalPath) Then Exit Function
    tempPath = SharedFolder() & "\tmp_" & fileName
    Set ts = f.CreateTextFile(tempPath, True, False)
    ts.Write text
    ts.Close
    f.MoveFile tempPath, finalPath
    WriteNewFile = True
End Function

' ---- Listing and deleting --------------------------------------------------

' Names of the .txt files starting with prefix ("block_" or "tx_"), sorted
' by name so blocks arrive in height order.
Public Function ListFileNames(ByVal prefix As String) As Collection
    Dim names As Collection
    Dim f As Object
    Dim nm As String
    Dim j As Long
    Dim placed As Boolean

    Set names = New Collection
    For Each f In Fso().GetFolder(SharedFolder()).Files
        nm = f.Name
        If LCase$(Left$(nm, Len(prefix))) = prefix And LCase$(Right$(nm, 4)) = ".txt" Then
            placed = False
            For j = 1 To names.Count
                If StrComp(CStr(names(j)), nm, vbBinaryCompare) > 0 Then
                    names.Add nm, , j
                    placed = True
                    Exit For
                End If
            Next j
            If Not placed Then names.Add nm
        End If
    Next f
    Set ListFileNames = names
End Function

' Deletes every block_ and tx_ file (and leftover tmp_ files). Returns how many.
Public Function DeleteAllFiles() As Long
    Dim kinds As Variant
    Dim names As Collection
    Dim k As Long
    Dim i As Long

    EnsureFolder
    kinds = Array("block_", "tx_", "tmp_")
    For k = LBound(kinds) To UBound(kinds)
        Set names = ListFileNames(CStr(kinds(k)))
        For i = 1 To names.Count
            Fso().DeleteFile SharedFolder() & "\" & CStr(names(i)), True
            DeleteAllFiles = DeleteAllFiles + 1
        Next i
    Next k
End Function

' ---- Reading ---------------------------------------------------------------

' Parses a tx_ file. Returns False with a reason when it is not valid.
Public Function ReadTxFile(ByVal fileName As String, ByRef tx As clsTransaction, ByRef reason As String) As Boolean
    Dim lines As Variant

    If Not ReadLines(fileName, lines) Then
        reason = modMessages.MSG_FILE_UNREADABLE
    ElseIf lines(0) <> modConstants.FILE_TX_HEADER Or UBound(lines) <> 1 Then
        reason = modMessages.MSG_FILE_HEADER
    Else
        ReadTxFile = ParseTxLine(CStr(lines(1)), tx, reason)
    End If
End Function

' Parses a block_ file: header, nine Key=Value lines in a fixed order, then
' one TX line per transaction. TxCount must match the number of TX lines.
Public Function ReadBlockFile(ByVal fileName As String, ByRef blk As clsBlock, ByRef reason As String) As Boolean
    Dim lines As Variant
    Dim b As clsBlock
    Dim tx As clsTransaction
    Dim i As Long

    If Not ReadLines(fileName, lines) Then
        reason = modMessages.MSG_FILE_UNREADABLE
        Exit Function
    End If
    If lines(0) <> modConstants.FILE_BLOCK_HEADER Or UBound(lines) < 9 Then
        reason = modMessages.MSG_FILE_HEADER
        Exit Function
    End If

    Set b = New clsBlock
    If Not ParseBlockFields(lines, b, reason) Then Exit Function
    If UBound(lines) - 9 <> b.TxCount Then
        reason = modMessages.FormatText(modMessages.MSG_FILE_TXCOUNT, b.TxCount, UBound(lines) - 9)
        Exit Function
    End If
    For i = 10 To UBound(lines)
        If Not ParseTxLine(CStr(lines(i)), tx, reason) Then
            reason = modMessages.FormatText(modMessages.MSG_FILE_TXLINE, i - 9) & ": " & reason
            Exit Function
        End If
        tx.BlockHeight = b.Height
        tx.Position = i - 10
        b.Transactions.Add tx
    Next i
    b.Source = modConstants.SRC_NETWORK
    Set blk = b
    ReadBlockFile = True
End Function

Private Function ParseBlockFields(ByRef lines As Variant, ByVal b As clsBlock, ByRef reason As String) As Boolean
    Dim keys As Variant
    Dim vals(1 To 9) As String
    Dim i As Long
    Dim heightV As Long
    Dim timeV As Long
    Dim nonceV As Long
    Dim bitsV As Long
    Dim countV As Long

    keys = Array("Height", "Timestamp", "PrevHash", "MerkleRoot", "Nonce", "DifficultyBits", "Miner", "Hash", "TxCount")
    For i = 1 To 9
        If Left$(CStr(lines(i)), Len(keys(i - 1)) + 1) <> keys(i - 1) & "=" Then
            reason = modMessages.FormatText(modMessages.MSG_FILE_FIELD, keys(i - 1))
            Exit Function
        End If
        vals(i) = Mid$(CStr(lines(i)), Len(keys(i - 1)) + 2)
        If Len(vals(i)) = 0 Then
            reason = modMessages.FormatText(modMessages.MSG_FILE_FIELD, keys(i - 1))
            Exit Function
        End If
    Next i

    ' Class fields cannot be passed ByRef (VBA hands over a copy), so use locals
    If Not NumberField(vals(1), heightV, "Height", reason) Then Exit Function
    If Not NumberField(vals(2), timeV, "Timestamp", reason) Then Exit Function
    If Not NumberField(vals(5), nonceV, "Nonce", reason) Then Exit Function
    If Not NumberField(vals(6), bitsV, "DifficultyBits", reason) Then Exit Function
    If Not NumberField(vals(9), countV, "TxCount", reason) Then Exit Function
    If Not IsHash(vals(3)) Or Not IsHash(vals(4)) Or Not IsHash(vals(8)) Then
        reason = modMessages.FormatText(modMessages.MSG_FILE_FIELD, "Hash")
        Exit Function
    End If
    b.Height = heightV
    b.Timestamp = timeV
    b.PrevHash = vals(3)
    b.MerkleRoot = vals(4)
    b.Nonce = nonceV
    b.DifficultyBits = bitsV
    b.Miner = vals(7)
    b.Hash = vals(8)
    b.TxCount = countV
    ParseBlockFields = True
End Function

Private Function NumberField(ByVal text As String, ByRef value As Long, ByVal fieldName As String, _
                             ByRef reason As String) As Boolean
    If modSheetUtils.TryParseLong(text, value) Then
        NumberField = True
    Else
        reason = modMessages.FormatText(modMessages.MSG_FILE_FIELD, fieldName)
    End If
End Function

Private Function IsHash(ByVal s As String) As Boolean
    IsHash = (Len(s) = 64) And Not (s Like "*[!0-9a-f]*")
End Function

' TX|TxID|FromAddress|FromPubKey|ToAddress|Amount|Fee|Nonce|Timestamp|Signature
' FromPubKey and Signature may be empty (coinbase).
Private Function ParseTxLine(ByVal txText As String, ByRef tx As clsTransaction, ByRef reason As String) As Boolean
    Dim p As Variant
    Dim t As clsTransaction
    Dim amountV As Long
    Dim feeV As Long
    Dim nonceV As Long
    Dim timeV As Long

    p = Split(txText, "|")
    If UBound(p) <> 9 Then
        reason = modMessages.FormatText(modMessages.MSG_FILE_FIELD, "TX")
        Exit Function
    End If
    If p(0) <> "TX" Or Len(p(1)) = 0 Or Len(p(2)) = 0 Or Len(p(4)) = 0 Then
        reason = modMessages.FormatText(modMessages.MSG_FILE_FIELD, "TX")
        Exit Function
    End If
    If Not NumberField(CStr(p(5)), amountV, "Amount", reason) Then Exit Function
    If Not NumberField(CStr(p(6)), feeV, "Fee", reason) Then Exit Function
    If Not NumberField(CStr(p(7)), nonceV, "Nonce", reason) Then Exit Function
    If Not NumberField(CStr(p(8)), timeV, "Timestamp", reason) Then Exit Function
    Set t = New clsTransaction
    t.TxID = p(1)
    t.FromAddress = p(2)
    t.FromPubKey = p(3)
    t.ToAddress = p(4)
    t.Amount = amountV
    t.Fee = feeV
    t.Nonce = nonceV
    t.Timestamp = timeV
    t.Signature = p(9)
    t.Source = modConstants.SRC_NETWORK
    Set tx = t
    ParseTxLine = True
End Function

' All lines of a file (CRLF or LF); one trailing empty line is dropped.
Private Function ReadLines(ByVal fileName As String, ByRef lines As Variant) As Boolean
    Dim ts As Object
    Dim text As String
    Dim parts() As String
    Dim n As Long

    On Error GoTo Failed
    Set ts = Fso().OpenTextFile(SharedFolder() & "\" & fileName, 1, False, 0)
    If Not ts.AtEndOfStream Then text = ts.ReadAll
    ts.Close
    If Len(text) = 0 Then Exit Function
    parts = Split(Replace(text, vbCrLf, vbLf), vbLf)
    n = UBound(parts)
    If Len(parts(n)) = 0 Then n = n - 1
    If n < 0 Then Exit Function
    ReDim Preserve parts(0 To n)
    lines = parts
    ReadLines = True
    Exit Function
Failed:
    ReadLines = False
End Function
