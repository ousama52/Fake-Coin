VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsResult"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Return value of every service call. Services never show a MsgBox;
' the entry macro shows Message.
Public Success As Boolean
Public Message As String
Public Count As Long

Public Sub Init(ByVal ok As Boolean, ByVal msg As String, Optional ByVal cnt As Long = 0)
    Success = ok
    Message = msg
    Count = cnt
End Sub
