VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBlock"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Plain data holder for one block. Timestamp is Unix seconds (UTC).
' Transactions holds the clsTransaction objects of the block in Merkle order
' (coinbase first) when they are loaded.
' CumulativeWork (Decimal) is always computed by the node from DifficultyBits;
' a value read from a table or file is never trusted.
Public Height As Long
Public Timestamp As Long
Public PrevHash As String
Public MerkleRoot As String
Public Nonce As Long
Public DifficultyBits As Long
Public Miner As String
Public TxCount As Long
Public Hash As String
Public CumulativeWork As Variant
Public Source As String
Public Transactions As Collection

Private Sub Class_Initialize()
    Set Transactions = New Collection
    CumulativeWork = CDec(0)
End Sub
