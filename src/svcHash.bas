Attribute VB_Name = "svcHash"
Option Explicit

' The ONLY place in the project that hashes.
' Engine 1: SHA-256 from .NET through COM (System.Security.Cryptography.SHA256Managed).
' That COM class needs the old .NET Framework 3.5 runtime. On a PC without it
' CreateObject fails with error 80131700, so engine 2 takes over: a plain VBA
' implementation of SHA-256. Both give identical output.
Private mSha As Object
Private mTried As Boolean
Private mUseCom As Boolean
Private mReady As Boolean
Private mK(0 To 63) As Long
Private mP2(0 To 31) As Long

' Lowercase hex SHA-256 of ASCII text. Callers pass plain decimal numbers
' (no separators) inside the text, exactly as DATA_TABLES.md writes the formulas.
Public Function Sha256(ByVal text As String) As String
    If Not mTried Then TryCreateCom
    ' An empty string gives COM an empty byte array, which some runtimes reject
    If mUseCom And Len(text) > 0 Then
        Sha256 = ShaCom(text)
    Else
        Sha256 = ShaVba(text)
    End If
End Function

' Number of leading zero BITS in a hex string: each leading "0" digit counts 4,
' then the first non-zero digit adds the zero bits in front of its top set bit.
' Examples: 0fff = 4, 0123 = 7, 1000 = 3, 8000 = 0, 00a0 = 8, 0000 = 16.
Public Function LeadingZeroBits(ByVal hexText As String) As Long
    Dim i As Long
    Dim v As Long

    For i = 1 To Len(hexText)
        v = InStr("0123456789abcdef", LCase$(Mid$(hexText, i, 1))) - 1
        If v < 0 Then Exit Function
        If v = 0 Then
            LeadingZeroBits = LeadingZeroBits + 4
        Else
            LeadingZeroBits = LeadingZeroBits + 4 - BitsNeeded(v)
            Exit Function
        End If
    Next i
End Function

' Bits needed to write v (1 to 15) in binary.
Private Function BitsNeeded(ByVal v As Long) As Long
    If v >= 8 Then
        BitsNeeded = 4
    ElseIf v >= 4 Then
        BitsNeeded = 3
    ElseIf v >= 2 Then
        BitsNeeded = 2
    Else
        BitsNeeded = 1
    End If
End Function

' "COM" or "VBA": which engine is in use (shown in the Log by Setup).
Public Function HashEngine() As String
    If Not mTried Then TryCreateCom
    If mUseCom Then
        HashEngine = "COM"
    Else
        HashEngine = "VBA"
    End If
End Function

Private Sub TryCreateCom()
    mTried = True
    On Error Resume Next
    Set mSha = CreateObject("System.Security.Cryptography.SHA256Managed")
    mUseCom = Not (mSha Is Nothing)
    Err.Clear
End Sub

Private Function ShaCom(ByVal text As String) As String
    Dim bytes() As Byte
    Dim hash As Variant
    Dim i As Long
    Dim s As String

    bytes = StrConv(text, vbFromUnicode)
    hash = mSha.ComputeHash_2(bytes)
    For i = LBound(hash) To UBound(hash)
        s = s & Right$("0" & LCase$(Hex$(hash(i))), 2)
    Next i
    ShaCom = s
End Function

' ---- Pure VBA SHA-256 ------------------------------------------------------
' VBA Long is a signed 32-bit number, so addition and shifts need helpers that
' behave like unsigned 32-bit arithmetic.

Private Sub InitTables()
    Dim k As Variant
    Dim i As Long

    mP2(0) = 1
    For i = 1 To 30
        mP2(i) = mP2(i - 1) * 2
    Next i
    mP2(31) = &H80000000
    k = Array(&H428A2F98, &H71374491, &HB5C0FBCF, &HE9B5DBA5, &H3956C25B, &H59F111F1, &H923F82A4, &HAB1C5ED5, _
              &HD807AA98, &H12835B01, &H243185BE, &H550C7DC3, &H72BE5D74, &H80DEB1FE, &H9BDC06A7, &HC19BF174, _
              &HE49B69C1, &HEFBE4786, &HFC19DC6, &H240CA1CC, &H2DE92C6F, &H4A7484AA, &H5CB0A9DC, &H76F988DA, _
              &H983E5152, &HA831C66D, &HB00327C8, &HBF597FC7, &HC6E00BF3, &HD5A79147, &H6CA6351, &H14292967, _
              &H27B70A85, &H2E1B2138, &H4D2C6DFC, &H53380D13, &H650A7354, &H766A0ABB, &H81C2C92E, &H92722C85, _
              &HA2BFE8A1, &HA81A664B, &HC24B8B70, &HC76C51A3, &HD192E819, &HD6990624, &HF40E3585, &H106AA070, _
              &H19A4C116, &H1E376C08, &H2748774C, &H34B0BCB5, &H391C0CB3, &H4ED8AA4A, &H5B9CCA4F, &H682E6FF3, _
              &H748F82EE, &H78A5636F, &H84C87814, &H8CC70208, &H90BEFFFA, &HA4506CEB, &HBEF9A3F7, &HC67178F2)
    For i = 0 To 63
        mK(i) = k(i)
    Next i
    mReady = True
End Sub

' (a + b) mod 2^32. The two 16-bit halves are added separately so the
' intermediate sums never overflow a signed Long.
Private Function AddU(ByVal a As Long, ByVal b As Long) As Long
    Dim lo As Long
    Dim hi As Long

    lo = (a And &HFFFF&) + (b And &HFFFF&)
    hi = (((a And &HFFFF0000) \ &H10000) And &HFFFF&) + (((b And &HFFFF0000) \ &H10000) And &HFFFF&) + (lo \ &H10000)
    AddU = ((hi And &H7FFF&) * &H10000) Or (lo And &HFFFF&)
    If (hi And &H8000&) <> 0 Then AddU = AddU Or &H80000000
End Function

' Logical shift right by n (1 to 30).
Private Function ShR(ByVal x As Long, ByVal n As Long) As Long
    If x >= 0 Then
        ShR = x \ mP2(n)
    Else
        ShR = ((x And &H7FFFFFFF) \ mP2(n)) Or mP2(31 - n)
    End If
End Function

' Shift left by n (1 to 30), dropping the bits that fall off the top.
Private Function ShL(ByVal x As Long, ByVal n As Long) As Long
    ShL = (x And (mP2(31 - n) - 1)) * mP2(n)
    If (x And mP2(31 - n)) <> 0 Then ShL = ShL Or &H80000000
End Function

Private Function RotR(ByVal x As Long, ByVal n As Long) As Long
    RotR = ShR(x, n) Or ShL(x, 32 - n)
End Function

' Big-endian 32-bit word from four bytes of the padded message.
Private Function WordAt(ByRef msg() As Byte, ByVal pos As Long) As Long
    WordAt = ((msg(pos) And &H7F&) * &H1000000) Or (msg(pos + 1) * &H10000) Or (msg(pos + 2) * &H100&) Or msg(pos + 3)
    If (msg(pos) And &H80&) <> 0 Then WordAt = WordAt Or &H80000000
End Function

' Message bytes + 0x80 + zeros + 64-bit big-endian bit length, a multiple of 64 bytes.
Private Function PadMessage(ByVal text As String) As Byte()
    Dim bytes() As Byte
    Dim msg() As Byte
    Dim n As Long
    Dim total As Long
    Dim bitLen As Long
    Dim i As Long

    If Len(text) > 0 Then
        bytes = StrConv(text, vbFromUnicode)
        n = UBound(bytes) + 1
    End If
    total = ((n + 8) \ 64 + 1) * 64
    ReDim msg(0 To total - 1)
    For i = 0 To n - 1
        msg(i) = bytes(i)
    Next i
    msg(n) = &H80
    bitLen = n * 8
    msg(total - 1) = bitLen And &HFF&
    msg(total - 2) = (bitLen \ &H100&) And &HFF&
    msg(total - 3) = (bitLen \ &H10000) And &HFF&
    msg(total - 4) = (bitLen \ &H1000000) And &HFF&
    PadMessage = msg
End Function

Private Function ShaVba(ByVal text As String) As String
    Dim msg() As Byte
    Dim w(0 To 63) As Long
    Dim hh(0 To 7) As Long
    Dim blk As Long
    Dim i As Long
    Dim s0 As Long
    Dim s1 As Long
    Dim ch As Long
    Dim maj As Long
    Dim t1 As Long
    Dim t2 As Long
    Dim va As Long, vb As Long, vc As Long, vd As Long
    Dim ve As Long, vf As Long, vg As Long, vh As Long
    Dim s As String

    If Not mReady Then InitTables
    msg = PadMessage(text)
    hh(0) = &H6A09E667: hh(1) = &HBB67AE85: hh(2) = &H3C6EF372: hh(3) = &HA54FF53A
    hh(4) = &H510E527F: hh(5) = &H9B05688C: hh(6) = &H1F83D9AB: hh(7) = &H5BE0CD19

    For blk = 0 To (UBound(msg) + 1) \ 64 - 1
        For i = 0 To 15
            w(i) = WordAt(msg, blk * 64 + i * 4)
        Next i
        For i = 16 To 63
            s0 = RotR(w(i - 15), 7) Xor RotR(w(i - 15), 18) Xor ShR(w(i - 15), 3)
            s1 = RotR(w(i - 2), 17) Xor RotR(w(i - 2), 19) Xor ShR(w(i - 2), 10)
            w(i) = AddU(AddU(w(i - 16), s0), AddU(w(i - 7), s1))
        Next i

        va = hh(0): vb = hh(1): vc = hh(2): vd = hh(3)
        ve = hh(4): vf = hh(5): vg = hh(6): vh = hh(7)
        For i = 0 To 63
            s1 = RotR(ve, 6) Xor RotR(ve, 11) Xor RotR(ve, 25)
            ch = (ve And vf) Xor ((Not ve) And vg)
            t1 = AddU(AddU(AddU(vh, s1), AddU(ch, mK(i))), w(i))
            s0 = RotR(va, 2) Xor RotR(va, 13) Xor RotR(va, 22)
            maj = (va And vb) Xor (va And vc) Xor (vb And vc)
            t2 = AddU(s0, maj)
            vh = vg: vg = vf: vf = ve
            ve = AddU(vd, t1)
            vd = vc: vc = vb: vb = va
            va = AddU(t1, t2)
        Next i
        hh(0) = AddU(hh(0), va): hh(1) = AddU(hh(1), vb): hh(2) = AddU(hh(2), vc): hh(3) = AddU(hh(3), vd)
        hh(4) = AddU(hh(4), ve): hh(5) = AddU(hh(5), vf): hh(6) = AddU(hh(6), vg): hh(7) = AddU(hh(7), vh)
    Next blk

    For i = 0 To 7
        s = s & LCase$(Right$("00000000" & Hex$(hh(i)), 8))
    Next i
    ShaVba = s
End Function
