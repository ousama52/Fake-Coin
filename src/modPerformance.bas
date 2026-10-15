Attribute VB_Name = "modPerformance"
Option Explicit

Private mDepth As Long
Private mScreenUpdating As Boolean
Private mEnableEvents As Boolean
Private mCalculation As XlCalculation

' Nested calls are counted so an inner TurnOn does not switch Excel back on
' while the outer macro is still running.
Public Sub TurnOff()
    If mDepth = 0 Then
        mScreenUpdating = Application.ScreenUpdating
        mEnableEvents = Application.EnableEvents
        mCalculation = Application.Calculation
    End If
    mDepth = mDepth + 1
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Public Sub TurnOn()
    If mDepth > 0 Then mDepth = mDepth - 1
    If mDepth > 0 Then Exit Sub
    Application.Calculation = mCalculation
    Application.EnableEvents = mEnableEvents
    Application.ScreenUpdating = mScreenUpdating
    Application.StatusBar = False
End Sub
