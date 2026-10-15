VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' PredeclaredId is True on purpose (allowed by the README rules): services call
' clsLogger.LogInfo without creating an object.

Public Sub LogInfo(ByVal procName As String, ByVal msg As String)
    WriteLine modConstants.LOG_INFO, procName, msg
End Sub

Public Sub LogWarning(ByVal procName As String, ByVal msg As String)
    WriteLine modConstants.LOG_WARNING, procName, msg
End Sub

Public Sub LogError(ByVal procName As String, ByVal msg As String)
    WriteLine modConstants.LOG_ERROR, procName, msg
End Sub

' Appends one row to tblLog.
Private Sub WriteLine(ByVal level As String, ByVal procName As String, ByVal msg As String)
    Dim rowData(1 To 1, 1 To 4) As Variant

    Debug.Print level & " | " & procName & " | " & msg
    ' A logging failure must never stop the real work
    On Error Resume Next
    rowData(1, 1) = modSheetUtils.NowText()
    rowData(1, 2) = level
    rowData(1, 3) = procName
    rowData(1, 4) = msg
    modSheetUtils.AppendRows modSheetUtils.GetTable(modConstants.TBL_LOG), rowData
End Sub
