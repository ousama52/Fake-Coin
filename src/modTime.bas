Attribute VB_Name = "modTime"
Option Explicit

' Unix time in UTC. VBA's Now is local time, so the system clock is read
' through the Windows API.
Private Type SYSTEMTIME
    wYear As Integer
    wMonth As Integer
    wDayOfWeek As Integer
    wDay As Integer
    wHour As Integer
    wMinute As Integer
    wSecond As Integer
    wMilliseconds As Integer
End Type

Private Declare PtrSafe Sub GetSystemTime Lib "kernel32" (ByRef lpSystemTime As SYSTEMTIME)

' Seconds since 1970-01-01 00:00:00 UTC.
Public Function UnixNowUtc() As Long
    Dim st As SYSTEMTIME
    Dim days As Long

    GetSystemTime st
    days = CLng(DateSerial(st.wYear, st.wMonth, st.wDay) - DateSerial(1970, 1, 1))
    UnixNowUtc = days * 86400 + CLng(st.wHour) * 3600 + CLng(st.wMinute) * 60 + CLng(st.wSecond)
End Function

' Readable text for logs and messages, for example 2026-01-01 00:00:00Z.
Public Function UnixToText(ByVal ts As Long) As String
    UnixToText = Format$(DateAdd("s", ts, DateSerial(1970, 1, 1)), "yyyy-mm-dd hh:nn:ss") & "Z"
End Function
