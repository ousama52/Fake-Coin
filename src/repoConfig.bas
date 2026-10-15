Attribute VB_Name = "repoConfig"
Option Explicit

' Value of a config key, "" when the key is missing or blank.
Private Function FindValue(ByVal key As String) As String
    Dim data As Variant
    Dim i As Long

    data = modSheetUtils.ReadTable(modSheetUtils.GetTable(modConstants.TBL_CONFIG))
    If IsEmpty(data) Then Exit Function
    For i = 1 To UBound(data, 1)
        If StrComp(modSheetUtils.CellText(data(i, 1)), key, vbTextCompare) = 0 Then
            FindValue = Trim$(modSheetUtils.CellText(data(i, 2)))
            Exit Function
        End If
    Next i
End Function

Public Function ConfigText(ByVal key As String, Optional ByVal defaultValue As String = "") As String
    ConfigText = FindValue(key)
    If Len(ConfigText) = 0 Then ConfigText = defaultValue
End Function

Public Function ConfigLong(ByVal key As String, ByVal defaultValue As Long) As Long
    Dim s As String

    s = FindValue(key)
    If Len(s) = 0 Then
        ConfigLong = defaultValue
    Else
        ConfigLong = modSheetUtils.CellLong(s)
    End If
End Function

' The NETWORK constants (and the local MempoolMaxTx) as a 1-based Long array,
' indexed with the PARAM_* constants. Every node must hold the same NETWORK values.
' Safe limits are applied here: DifficultyInterval is at least 3, the bit limits
' stay between 1 and 50 (2 ^ bits must stay exact), HalvingInterval is at least 1.
Public Function LoadChainParams() As Variant
    Dim p() As Long

    ReDim p(1 To modConstants.PARAM_COUNT)
    p(modConstants.PARAM_REWARD) = ConfigLong(modConstants.CFG_BLOCK_REWARD, 50)
    p(modConstants.PARAM_HALVING) = ConfigLong(modConstants.CFG_HALVING, 10)
    p(modConstants.PARAM_MAXTX) = ConfigLong(modConstants.CFG_MAX_TX, 5)
    p(modConstants.PARAM_MINFEE) = ConfigLong(modConstants.CFG_MIN_FEE, 1)
    p(modConstants.PARAM_INIT_BITS) = ConfigLong(modConstants.CFG_INIT_BITS, 12)
    p(modConstants.PARAM_TARGET_SECONDS) = ConfigLong(modConstants.CFG_TARGET_SECONDS, 30)
    p(modConstants.PARAM_DIFF_INTERVAL) = ConfigLong(modConstants.CFG_DIFF_INTERVAL, 10)
    p(modConstants.PARAM_MIN_BITS) = ConfigLong(modConstants.CFG_MIN_BITS, 4)
    p(modConstants.PARAM_MAX_BITS) = ConfigLong(modConstants.CFG_MAX_BITS, 32)
    p(modConstants.PARAM_MATURITY) = ConfigLong(modConstants.CFG_MATURITY, 3)
    p(modConstants.PARAM_MAX_FUTURE) = ConfigLong(modConstants.CFG_MAX_FUTURE, 600)
    p(modConstants.PARAM_MEMPOOL_MAX) = ConfigLong(modConstants.CFG_MEMPOOL_MAX, 50)

    If p(modConstants.PARAM_HALVING) < 1 Then p(modConstants.PARAM_HALVING) = 1
    If p(modConstants.PARAM_DIFF_INTERVAL) < 3 Then p(modConstants.PARAM_DIFF_INTERVAL) = 3
    If p(modConstants.PARAM_TARGET_SECONDS) < 1 Then p(modConstants.PARAM_TARGET_SECONDS) = 1
    If p(modConstants.PARAM_MIN_BITS) < 1 Then p(modConstants.PARAM_MIN_BITS) = 1
    If p(modConstants.PARAM_MAX_BITS) > 50 Then p(modConstants.PARAM_MAX_BITS) = 50
    If p(modConstants.PARAM_MAX_BITS) < p(modConstants.PARAM_MIN_BITS) Then
        p(modConstants.PARAM_MAX_BITS) = p(modConstants.PARAM_MIN_BITS)
    End If
    If p(modConstants.PARAM_MATURITY) < 0 Then p(modConstants.PARAM_MATURITY) = 0
    If p(modConstants.PARAM_MEMPOOL_MAX) < 1 Then p(modConstants.PARAM_MEMPOOL_MAX) = 1
    LoadChainParams = p
End Function

Public Function GetMinerFromControl() As String
    GetMinerFromControl = modSheetUtils.ControlValue(modConstants.CTL_MINER)
End Function
