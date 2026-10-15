Attribute VB_Name = "modMain"
Option Explicit

' FAKE - FOR LEARNING ONLY.
' Entry macros only: set up error handling, call one service, show the result.
' The error text is copied before LogError runs because logging can overwrite
' the Err object.

Public Sub RunSetup()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = modSetup.EnsureAll()
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunSetup", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunCreateWallet()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcWallet.CreateWallet(modSheetUtils.ControlValue(modConstants.CTL_WALLET))
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunCreateWallet", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunSendCoins()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcTransaction.SendCoins(modSheetUtils.ControlValue(modConstants.CTL_FROM), _
                                          modSheetUtils.ControlValue(modConstants.CTL_TO), _
                                          modSheetUtils.ControlValue(modConstants.CTL_AMOUNT), _
                                          modSheetUtils.ControlValue(modConstants.CTL_FEE), _
                                          modSheetUtils.ControlValue(modConstants.CTL_NONCE))
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunSendCoins", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunMineBlock()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcMining.MineBlock(repoConfig.GetMinerFromControl())
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunMineBlock", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunSync()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcConsensus.Sync()
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunSync", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunValidateChain()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcValidation.ValidateChain()
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunValidateChain", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunRebuildFromFolder()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcConsensus.RebuildFromFolder()
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunRebuildFromFolder", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunTamperTest()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcValidation.TamperBlock(modSheetUtils.ControlValue(modConstants.CTL_HEIGHT))
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunTamperTest", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunResetLocal()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    If MsgBox(modMessages.MSG_CONFIRM_RESET_LOCAL, vbYesNo + vbQuestion) <> vbYes Then
        MsgBox modMessages.MSG_CANCELLED, vbInformation
        Exit Sub
    End If
    modPerformance.TurnOff
    Set result = svcConsensus.ResetLocalChain()
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunResetLocal", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunResetNetwork()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    If MsgBox(modMessages.MSG_CONFIRM_RESET_NET, vbYesNo + vbExclamation) <> vbYes Then
        MsgBox modMessages.MSG_CANCELLED, vbInformation
        Exit Sub
    End If
    modPerformance.TurnOff
    Set result = svcConsensus.ResetNetwork()
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunResetNetwork", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

Public Sub RunTxStatus()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcTransaction.TxStatus(modSheetUtils.ControlValue(modConstants.CTL_TXID))
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunTxStatus", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

' Attack simulation: C4 victim (wallet name or address), C5 receiver, C6 amount, C7 fee.
Public Sub RunForgeTest()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcTransaction.ForgeTransaction(modSheetUtils.ControlValue(modConstants.CTL_FROM), _
                                                 modSheetUtils.ControlValue(modConstants.CTL_TO), _
                                                 modSheetUtils.ControlValue(modConstants.CTL_AMOUNT), _
                                                 modSheetUtils.ControlValue(modConstants.CTL_FEE))
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunForgeTest", Err
    MsgBox modMessages.FormatText(modMessages.MSG_ERROR, errText), vbCritical
    Resume Cleanup
End Sub

' Records the message as the Control sheet's last status and shows it.

Private Sub ShowResult(ByVal result As clsResult)
    Dim icon As VbMsgBoxStyle

    modSheetUtils.SetStatus result.Message
    If result.Success Then
        icon = vbInformation
    Else
        icon = vbExclamation
    End If
    MsgBox result.Message, icon
End Sub
