Attribute VB_Name = "svcCrypto"
Option Explicit

' ============================================================================
' FOR LEARNING ONLY. RSA-1024 signing with SHA-256 (PKCS#1 v1.5), the same
' scheme as .NET SignData(..., "SHA256"). Nothing else in the project calls RSA.
'
' The .NET COM class RSACryptoServiceProvider would be the obvious choice, but that
' class needs the old .NET Framework 3.5 runtime and fails with error 80131700
' on a PC without it, so the same RSA operations go through the Windows
' CryptoAPI (advapi32), which every Windows PC has.
'
' Formats:
'   PublicKey  = base64 of the 128 byte big-endian modulus (exponent is 65537)
'   Signature  = base64 of the 128 byte big-endian RSA signature
'   Private key text = base64 of the exported private key blob (kept only in a key file)
'   Address    = "FC" & first 40 hex characters of SHA256(PublicKey)
' ============================================================================

Private Declare PtrSafe Function CryptAcquireContextA Lib "advapi32.dll" (ByRef phProv As LongPtr, _
    ByVal pszContainer As String, ByVal pszProvider As String, ByVal dwProvType As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptReleaseContext Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptGenKey Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal algId As Long, _
    ByVal dwFlags As Long, ByRef phKey As LongPtr) As Long
Private Declare PtrSafe Function CryptDestroyKey Lib "advapi32.dll" (ByVal hKey As LongPtr) As Long
Private Declare PtrSafe Function CryptExportKey Lib "advapi32.dll" (ByVal hKey As LongPtr, ByVal hExpKey As LongPtr, _
    ByVal dwBlobType As Long, ByVal dwFlags As Long, ByRef pbData As Any, ByRef pdwDataLen As Long) As Long
Private Declare PtrSafe Function CryptImportKey Lib "advapi32.dll" (ByVal hProv As LongPtr, ByRef pbData As Any, _
    ByVal dwDataLen As Long, ByVal hPubKey As LongPtr, ByVal dwFlags As Long, ByRef phKey As LongPtr) As Long
Private Declare PtrSafe Function CryptCreateHash Lib "advapi32.dll" (ByVal hProv As LongPtr, ByVal algId As Long, _
    ByVal hKey As LongPtr, ByVal dwFlags As Long, ByRef phHash As LongPtr) As Long
Private Declare PtrSafe Function CryptHashData Lib "advapi32.dll" (ByVal hHash As LongPtr, ByRef pbData As Any, _
    ByVal dwDataLen As Long, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptSignHashA Lib "advapi32.dll" (ByVal hHash As LongPtr, ByVal dwKeySpec As Long, _
    ByVal sDescription As LongPtr, ByVal dwFlags As Long, ByRef pbSignature As Any, ByRef pdwSigLen As Long) As Long
Private Declare PtrSafe Function CryptVerifySignatureA Lib "advapi32.dll" (ByVal hHash As LongPtr, ByRef pbSignature As Any, _
    ByVal dwSigLen As Long, ByVal hPubKey As LongPtr, ByVal sDescription As LongPtr, ByVal dwFlags As Long) As Long
Private Declare PtrSafe Function CryptDestroyHash Lib "advapi32.dll" (ByVal hHash As LongPtr) As Long

Private Const PROV_RSA_AES As Long = 24
Private Const CRYPT_VERIFYCONTEXT As Long = &HF0000000
Private Const AT_SIGNATURE As Long = 2
Private Const CRYPT_EXPORTABLE As Long = 1
Private Const PRIVATEKEYBLOB As Long = 7
Private Const PUBLICKEYBLOB As Long = 6
Private Const CALG_SHA_256 As Long = &H800C&  ' trailing & : a 4-digit hex literal is a signed Integer without it
Private Const KEY_BITS As Long = 1024
Private Const MODULUS_BYTES As Long = 128
Private Const BLOB_PREFIX As Long = 20
Private Const PROVIDER_NAME As String = "Microsoft Enhanced RSA and AES Cryptographic Provider"

Private mDoc As Object
Private mVerified As Object

' New RSA-1024 key pair. privateKeyText goes only into a key file; publicKey is
' the text used in tables and files.
Public Sub GenerateKeyPair(ByRef privateKeyText As String, ByRef publicKey As String)
    Dim hProv As LongPtr
    Dim hKey As LongPtr
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String

    On Error GoTo Fail
    hProv = OpenProvider()
    CheckApi CryptGenKey(hProv, AT_SIGNATURE, (KEY_BITS * &H10000) Or CRYPT_EXPORTABLE, hKey), "CryptGenKey"
    privateKeyText = ToBase64(ExportBlob(hKey, PRIVATEKEYBLOB))
    publicKey = ModulusFromPublicBlob(ExportBlob(hKey, PUBLICKEYBLOB))
Done:
    On Error Resume Next
    If hKey <> 0 Then CryptDestroyKey hKey
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, errSrc, errDesc
    Exit Sub
Fail:
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    Resume Done
End Sub

' "FC" + first 40 hex characters of SHA256(PublicKey as ASCII text).
Public Function AddressFromPublicKey(ByVal publicKey As String) As String
    AddressFromPublicKey = modConstants.ADDRESS_PREFIX & Left$(svcHash.Sha256(publicKey), 40)
End Function

' RSA signature over the ASCII bytes of text (SHA-256), as base64.
Public Function Sign(ByVal privateKeyText As String, ByVal text As String) As String
    Dim hProv As LongPtr
    Dim hKey As LongPtr
    Dim hHash As LongPtr
    Dim blob() As Byte
    Dim data() As Byte
    Dim sig() As Byte
    Dim sigLen As Long
    Dim errNum As Long
    Dim errSrc As String
    Dim errDesc As String

    On Error GoTo Fail
    hProv = OpenProvider()
    blob = FromBase64(privateKeyText)
    CheckApi CryptImportKey(hProv, blob(0), UBound(blob) + 1, 0, 0, hKey), "CryptImportKey"
    CheckApi CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash), "CryptCreateHash"
    data = StrConv(text, vbFromUnicode)
    CheckApi CryptHashData(hHash, data(0), UBound(data) + 1, 0), "CryptHashData"
    CheckApi CryptSignHashA(hHash, AT_SIGNATURE, 0, 0, ByVal 0&, sigLen), "CryptSignHash(size)"
    ReDim sig(0 To sigLen - 1)
    CheckApi CryptSignHashA(hHash, AT_SIGNATURE, 0, 0, sig(0), sigLen), "CryptSignHash"
    Sign = ToBase64(ReverseBytes(sig, 0, sigLen))
Done:
    On Error Resume Next
    If hHash <> 0 Then CryptDestroyHash hHash
    If hKey <> 0 Then CryptDestroyKey hKey
    If hProv <> 0 Then CryptReleaseContext hProv, 0
    On Error GoTo 0
    If errNum <> 0 Then Err.Raise errNum, errSrc, errDesc
    Exit Function
Fail:
    errNum = Err.Number
    errSrc = Err.Source
    errDesc = Err.Description
    Resume Done
End Function

' True only when signatureBase64 is a valid RSA signature of text under the
' public key. Never raises: any failure (bad base64, wrong length) is False.
' Successful checks are remembered so a chain replay does not repeat them.
Public Function VerifySignature(ByVal publicKey As String, ByVal text As String, _
                                ByVal signatureBase64 As String) As Boolean
    Dim cacheKey As String
    Dim hProv As LongPtr
    Dim hPub As LongPtr
    Dim hHash As LongPtr
    Dim blob() As Byte
    Dim sig() As Byte
    Dim data() As Byte
    Dim raw() As Byte

    cacheKey = publicKey & "|" & text & "|" & signatureBase64
    If mVerified Is Nothing Then Set mVerified = CreateObject("Scripting.Dictionary")
    If mVerified.Exists(cacheKey) Then
        VerifySignature = True
        Exit Function
    End If

    On Error GoTo Done
    raw = FromBase64(signatureBase64)
    If UBound(raw) + 1 <> MODULUS_BYTES Then GoTo Done
    sig = ReverseBytes(raw, 0, MODULUS_BYTES)
    blob = BuildPublicBlob(publicKey)
    hProv = OpenProvider()
    If CryptImportKey(hProv, blob(0), UBound(blob) + 1, 0, 0, hPub) = 0 Then GoTo Done
    If CryptCreateHash(hProv, CALG_SHA_256, 0, 0, hHash) = 0 Then GoTo Done
    data = StrConv(text, vbFromUnicode)
    If CryptHashData(hHash, data(0), UBound(data) + 1, 0) = 0 Then GoTo Done
    VerifySignature = (CryptVerifySignatureA(hHash, sig(0), MODULUS_BYTES, hPub, 0, 0) <> 0)
    If VerifySignature Then mVerified(cacheKey) = True
Done:
    On Error Resume Next
    If hHash <> 0 Then CryptDestroyHash hHash
    If hPub <> 0 Then CryptDestroyKey hPub
    If hProv <> 0 Then CryptReleaseContext hProv, 0
End Function

' ---- Helpers ---------------------------------------------------------------

Private Function OpenProvider() As LongPtr
    Dim h As LongPtr

    CheckApi CryptAcquireContextA(h, vbNullString, PROVIDER_NAME, PROV_RSA_AES, CRYPT_VERIFYCONTEXT), "CryptAcquireContext"
    OpenProvider = h
End Function

Private Sub CheckApi(ByVal ok As Long, ByVal what As String)
    Dim code As Long

    If ok <> 0 Then Exit Sub
    code = Err.LastDllError
    Err.Raise modConstants.ERR_APP + 60, "svcCrypto." & what, _
              modMessages.FormatText(modMessages.MSG_CRYPTO_FAIL, what, code)
End Sub

Private Function ExportBlob(ByVal hKey As LongPtr, ByVal blobType As Long) As Byte()
    Dim blob() As Byte
    Dim n As Long

    CheckApi CryptExportKey(hKey, 0, blobType, 0, ByVal 0&, n), "CryptExportKey(size)"
    ReDim blob(0 To n - 1)
    CheckApi CryptExportKey(hKey, 0, blobType, 0, blob(0), n), "CryptExportKey"
    ExportBlob = blob
End Function

' Public key text from an exported PUBLICKEYBLOB: the modulus, big-endian, base64.
' The blob stores it little-endian after a 20 byte header.
Private Function ModulusFromPublicBlob(ByRef blob() As Byte) As String
    Dim exponent As Long

    exponent = CLng(blob(16)) + CLng(blob(17)) * 256 + CLng(blob(18)) * 65536
    If exponent <> 65537 Then
        Err.Raise modConstants.ERR_APP + 61, "svcCrypto.ModulusFromPublicBlob", modMessages.MSG_CRYPTO_EXPONENT
    End If
    ModulusFromPublicBlob = ToBase64(ReverseBytes(blob, BLOB_PREFIX, MODULUS_BYTES))
End Function

' PUBLICKEYBLOB for CryptImportKey: header, RSA1 magic, 1024 bits, exponent
' 65537, then the modulus little-endian.
Private Function BuildPublicBlob(ByVal publicKey As String) As Byte()
    Dim modulus() As Byte
    Dim blob() As Byte
    Dim head As Variant
    Dim i As Long

    modulus = FromBase64(publicKey)
    If UBound(modulus) + 1 <> MODULUS_BYTES Then
        Err.Raise modConstants.ERR_APP + 62, "svcCrypto.BuildPublicBlob", modMessages.MSG_CRYPTO_BAD_KEY
    End If
    head = Array(6, 2, 0, 0, &H0, &H24, 0, 0, &H52, &H53, &H41, &H31, 0, 4, 0, 0, 1, 0, 1, 0)
    ReDim blob(0 To BLOB_PREFIX + MODULUS_BYTES - 1)
    For i = 0 To BLOB_PREFIX - 1
        blob(i) = head(i)
    Next i
    For i = 0 To MODULUS_BYTES - 1
        blob(BLOB_PREFIX + i) = modulus(MODULUS_BYTES - 1 - i)
    Next i
    BuildPublicBlob = blob
End Function

' count bytes of src starting at index first, in reverse order.
Private Function ReverseBytes(ByRef src() As Byte, ByVal first As Long, ByVal count As Long) As Byte()
    Dim result() As Byte
    Dim i As Long

    ReDim result(0 To count - 1)
    For i = 0 To count - 1
        result(i) = src(first + count - 1 - i)
    Next i
    ReverseBytes = result
End Function

' Base64 through MSXML. MSXML breaks long base64 text into lines; remove them.
Private Function ToBase64(ByRef bytes() As Byte) As String
    Dim node As Object

    Set node = NewBase64Node()
    node.nodeTypedValue = bytes
    ToBase64 = Replace(Replace(CStr(node.Text), vbLf, ""), vbCr, "")
End Function

Private Function FromBase64(ByVal text As String) As Byte()
    Dim node As Object

    Set node = NewBase64Node()
    node.Text = text
    FromBase64 = node.nodeTypedValue
End Function

Private Function NewBase64Node() As Object
    Dim node As Object

    If mDoc Is Nothing Then Set mDoc = CreateObject("MSXML2.DOMDocument")
    Set node = mDoc.createElement("b64")
    node.DataType = "bin.base64"
    Set NewBase64Node = node
End Function
