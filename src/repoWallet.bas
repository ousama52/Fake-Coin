Attribute VB_Name = "repoWallet"
Option Explicit

' tblWallets columns (public so svcWallet can read the array).
' The private key is NOT stored here, only the path of its key file.
Public Const WLC_NAME As Long = 1
Public Const WLC_ADDRESS As Long = 2
Public Const WLC_PUBKEY As Long = 3
Public Const WLC_KEYFILE As Long = 4
Public Const WLC_CREATED As Long = 5

' 1-based 2D array, Empty when there are no wallets. Test IsEmpty before use.
Public Function LoadWallets() As Variant
    LoadWallets = modSheetUtils.ReadTable(modSheetUtils.GetTable(modConstants.TBL_WALLETS))
End Function

Public Sub AppendWallet(ByVal walletName As String, ByVal address As String, ByVal publicKey As String, _
                        ByVal keyFile As String, ByVal createdAt As Long)
    Dim data(1 To 1, 1 To 5) As Variant
    Dim rows As Variant

    data(1, WLC_NAME) = walletName
    data(1, WLC_ADDRESS) = address
    data(1, WLC_PUBKEY) = publicKey
    data(1, WLC_KEYFILE) = keyFile
    data(1, WLC_CREATED) = createdAt
    rows = data
    modSheetUtils.AppendRows modSheetUtils.GetTable(modConstants.TBL_WALLETS), rows
End Sub

' Row of a wallet name (case-insensitive), 0 when not found.
Public Function FindWalletByName(ByRef data As Variant, ByVal walletName As String) As Long
    Dim i As Long

    If modSheetUtils.RowCount(data) = 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(modSheetUtils.CellText(data(i, WLC_NAME)), walletName, vbTextCompare) = 0 Then
            FindWalletByName = i
            Exit Function
        End If
    Next i
End Function

' Row of an address, 0 when not found.
Public Function FindWalletByAddress(ByRef data As Variant, ByVal address As String) As Long
    Dim i As Long

    If modSheetUtils.RowCount(data) = 0 Then Exit Function
    For i = 1 To UBound(data, 1)
        If modSheetUtils.CellText(data(i, WLC_ADDRESS)) = address Then
            FindWalletByAddress = i
            Exit Function
        End If
    Next i
End Function
