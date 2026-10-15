Attribute VB_Name = "repoBalance"
Option Explicit

' tblBalances columns
Public Const BAC_ADDRESS As Long = 1
Public Const BAC_NAME As Long = 2
Public Const BAC_BALANCE As Long = 3
Public Const BAC_SPENDABLE As Long = 4
Public Const BAC_NONCE As Long = 5
Public Const BAC_COUNT As Long = 5

' 1-based 2D array, Empty when the table has no data.
Public Function LoadBalances() As Variant
    LoadBalances = modSheetUtils.ReadTable(modSheetUtils.GetTable(modConstants.TBL_BALANCES))
End Function

' Replaces the whole table. Balances are always rebuilt, never edited.
Public Sub SaveBalances(ByRef data As Variant)
    modSheetUtils.WriteTable modSheetUtils.GetTable(modConstants.TBL_BALANCES), data
End Sub
