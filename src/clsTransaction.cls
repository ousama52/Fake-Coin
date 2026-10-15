VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTransaction"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Plain data holder for one transaction (Mempool row or Transactions row).
' Timestamp is Unix seconds (UTC). BlockHeight and Position are only meaningful
' for confirmed transactions; Source only for Mempool rows.
' For a coinbase, FromAddress is COINBASE and FromPubKey and Signature are empty.
Public TxID As String
Public BlockHeight As Long
Public Position As Long
Public FromAddress As String
Public FromPubKey As String
Public ToAddress As String
Public Amount As Long
Public Fee As Long
Public Nonce As Long
Public Timestamp As Long
Public Signature As String
Public Source As String
