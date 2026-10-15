Attribute VB_Name = "svcDifficulty"
Option Explicit

' Required difficulty (in BITS) for a block at the given height, computed from
' the chain itself, never taken from Config (except the fixed genesis value).
' chain holds the blocks before the new one; index = Height + 1.
' The adjustment uses integers only.
Public Function RequiredBits(ByVal height As Long, ByVal chain As Collection, ByRef params As Variant) As Long
    Dim parentBits As Long
    Dim interval As Long

    If height <= 0 Then
        RequiredBits = modConstants.GENESIS_BITS
        Exit Function
    End If
    parentBits = chain(height).DifficultyBits
    interval = params(modConstants.PARAM_DIFF_INTERVAL)
    If height Mod interval <> 0 Then
        RequiredBits = parentBits
    Else
        RequiredBits = Adjusted(height, chain, params, parentBits)
    End If
End Function

' Adjustment block: compare how long the last blocks really took with how
' long they should have taken (TargetBlockSeconds per gap).
Private Function Adjusted(ByVal height As Long, ByVal chain As Collection, ByRef params As Variant, _
                          ByVal parentBits As Long) As Long
    Dim startH As Long
    Dim endH As Long
    Dim expected As Variant
    Dim actual As Variant
    Dim b As Long

    startH = height - params(modConstants.PARAM_DIFF_INTERVAL)
    If startH = 0 Then startH = 1
    endH = height - 1
    expected = CDec(params(modConstants.PARAM_TARGET_SECONDS)) * CDec(endH - startH)
    actual = CDec(chain(endH + 1).Timestamp) - CDec(chain(startH + 1).Timestamp)
    If actual < 1 Then actual = CDec(1)

    b = parentBits
    If actual * 4 <= expected Then
        b = b + 2
    ElseIf actual * 2 <= expected Then
        b = b + 1
    ElseIf actual >= expected * 4 Then
        b = b - 2
    ElseIf actual >= expected * 2 Then
        b = b - 1
    End If
    If b < params(modConstants.PARAM_MIN_BITS) Then b = params(modConstants.PARAM_MIN_BITS)
    If b > params(modConstants.PARAM_MAX_BITS) Then b = params(modConstants.PARAM_MAX_BITS)
    Adjusted = b
End Function

' Work of one block = 2 ^ DifficultyBits, as a Decimal.
Public Function WorkOfBits(ByVal bits As Long) As Variant
    WorkOfBits = CDec(2 ^ bits)
End Function

' Sum of the work of every block in the chain, always computed here from
' DifficultyBits, never read from a table or file.
Public Function ChainWork(ByVal chain As Collection) As Variant
    Dim total As Variant
    Dim i As Long

    total = CDec(0)
    For i = 1 To chain.Count
        total = total + WorkOfBits(chain(i).DifficultyBits)
    Next i
    ChainWork = total
End Function

' Fills every block's CumulativeWork (running sum) so it can be saved.
Public Sub AssignCumulativeWork(ByVal chain As Collection)
    Dim total As Variant
    Dim i As Long

    total = CDec(0)
    For i = 1 To chain.Count
        total = total + WorkOfBits(chain(i).DifficultyBits)
        chain(i).CumulativeWork = total
    Next i
End Sub
