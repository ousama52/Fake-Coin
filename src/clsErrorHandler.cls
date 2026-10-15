VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsErrorHandler"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' PredeclaredId is True on purpose (allowed by the README rules): entry macros call
' clsErrorHandler.LogError "RunMineBlock", Err without creating an object.

Public Sub LogError(ByVal procName As String, ByVal errObj As ErrObject)
    Dim msg As String

    ' Read Err first: any On Error statement afterwards clears it
    msg = FormatError(errObj)
    On Error Resume Next
    clsLogger.LogError procName, msg
End Sub

Public Function FormatError(ByVal errObj As ErrObject) As String
    FormatError = "Err " & errObj.Number & ": " & errObj.Description & " [" & errObj.Source & "]"
End Function
