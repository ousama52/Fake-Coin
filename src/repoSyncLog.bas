Attribute VB_Name = "repoSyncLog"
Option Explicit

' tblSyncLog columns
Public Const SLC_FILE As Long = 1
Public Const SLC_KIND As Long = 2
Public Const SLC_ITEM As Long = 3
Public Const SLC_RESULT As Long = 4
Public Const SLC_REASON As Long = 5
Public Const SLC_TIME As Long = 6
Public Const SLC_COUNT As Long = 6

' FileName -> last Result, as a Scripting.Dictionary. Sync skips every file
' that is in here unless its last result was Orphan (those are retried).
Public Function LoadSeenFiles() As Object
    Dim map As Object
    Dim data As Variant
    Dim i As Long

    Set map = CreateObject("Scripting.Dictionary")
    data = modSheetUtils.ReadTable(modSheetUtils.GetTable(modConstants.TBL_SYNCLOG))
    If Not IsEmpty(data) Then
        For i = 1 To UBound(data, 1)
            map(modSheetUtils.CellText(data(i, SLC_FILE))) = modSheetUtils.CellText(data(i, SLC_RESULT))
        Next i
    End If
    Set LoadSeenFiles = map
End Function

' Adds rows (1-based 2D array with the 6 columns) below the existing ones.
Public Sub AppendSyncLog(ByRef rows As Variant)
    modSheetUtils.AppendRows modSheetUtils.GetTable(modConstants.TBL_SYNCLOG), rows
End Sub

Public Sub ClearSyncLog()
    Dim none As Variant

    modSheetUtils.WriteTable modSheetUtils.GetTable(modConstants.TBL_SYNCLOG), none
End Sub
