# DATA_TABLES.md - FakeCoin

FOR LEARNING ONLY. This file is the schema of the tables, the exact formulas and the file formats.
Every node must follow it byte for byte.

Each table is an Excel Table (ListObject) on its own worksheet, starting at A1 (headers in row 1).
Sheet name = table name without the `tbl` prefix. Example: sheet `Blocks` holds table `tblBlocks`.
Column names are the exact header text. Keep this order.

Type legend: Text, Int (whole number), Enum (fixed list, validated).
All hashes are 64 character lowercase hex strings, stored as Text.
Format hash, key, address, signature and similar columns as Text (`@`) so Excel never changes them.
A table with no data keeps one blank row; a row whose first column is empty counts as "no data".

## Table relationships

```
Blocks 1---* Transactions          (Transactions.BlockHeight = Blocks.Height)
Wallets / Balances / Transactions  (linked through the address)
Mempool -> mined into -> Blocks + Transactions
SyncLog records every file read from the shared folder
```

---

## tblConfig (sheet: Config)

Key/Value table, values stored as text. `Key` is unique. NETWORK values must be identical on every node.

| Key | Default | Meaning |
|---|---|---|
| CoinName | `FakeCoin` | Display only |
| NodeID | `NodeA` | Name of this node. Change it in each copy of the workbook |
| SharedFolder | `C:\FakeCoin\network` | Folder shared by all nodes. Must exist |
| WalletFolder | `C:\FakeCoin\wallets` | Local. Key files go to `<WalletFolder>\<NodeID>\<name>.key` |
| BlockReward | 50 | NETWORK. Reward for the first halving era |
| HalvingInterval | 10 | NETWORK. Blocks per halving era |
| MaxTxPerBlock | 5 | NETWORK. Maximum non-coinbase transactions per block |
| MinFee | 1 | NETWORK. Minimum fee per transaction |
| InitialDifficultyBits | 12 | NETWORK. Documents the genesis difficulty (the genesis block itself is fixed) |
| TargetBlockSeconds | 30 | NETWORK |
| DifficultyInterval | 10 | NETWORK. Must be >= 3 (smaller values are treated as 3) |
| MinDifficultyBits | 4 | NETWORK |
| MaxDifficultyBits | 32 | NETWORK (values above 50 are treated as 50) |
| CoinbaseMaturity | 3 | NETWORK. 0 turns maturity off |
| MaxFutureSeconds | 600 | NETWORK |
| MempoolMaxTx | 50 | Local |

## tblWallets (sheet: Wallets)

This node's own wallets only. NO secret is stored here.

| Column | Type | Rules |
|---|---|---|
| Name | Text | Required, unique on this node. Letters, digits and `_` only |
| Address | Text | `FC` + first 40 hex of `SHA256(PublicKey)` (42 characters) |
| PublicKey | Text | Base64 of the 128 byte big-endian RSA modulus (exponent is always 65537) |
| KeyFile | Text | Path of the private key file |
| CreatedAt | Int | Unix seconds (UTC) |

Key file: ASCII, CRLF. Line 1 `FAKECOIN-KEY`, line 2 the private key as base64 (an exported
Windows CryptoAPI private key blob). It is never written to the workbook or the shared folder and is
NOT encrypted. Existing key files are never overwritten.

## tblMempool (sheet: Mempool)

| Column | Type | Rules |
|---|---|---|
| TxID | Text | Key |
| FromAddress | Text | Sender address |
| FromPubKey | Text | Sender public key |
| ToAddress | Text | `FC` + 40 hex |
| Amount | Int | > 0 |
| Fee | Int | >= Config.MinFee |
| Nonce | Int | Sender's transaction counter: 0 for the first, then 1, 2, ... (pending ones count) |
| Timestamp | Int | Unix seconds (UTC) |
| Signature | Text | Base64 RSA signature of the TxID text |
| Source | Enum | `Local`, `Network` |

## tblBlocks (sheet: Blocks)

The local copy of the best chain, in height order. Height 0 is genesis.

| Column | Type | Rules |
|---|---|---|
| Height | Int | Key. 0, 1, 2, ... without gaps |
| Timestamp | Int | Unix seconds (UTC). Validated |
| PrevHash | Text | Hash of the previous block. 64 zeros for genesis |
| MerkleRoot | Text | See formulas |
| Nonce | Int | Proof-of-work counter |
| DifficultyBits | Int | Must equal the value computed from the chain |
| Miner | Text | Miner ADDRESS (any address), or `GENESIS` |
| TxCount | Int | Number of transactions in the block including the coinbase |
| Hash | Text | See formulas. Needs `DifficultyBits` leading zero bits (except genesis) |
| CumulativeWork | Int | Sum of `2 ^ DifficultyBits` up to this block. Always computed by the node, never trusted |
| Source | Enum | `Genesis`, `Mined`, `Network` |

## tblTransactions (sheet: Transactions)

Confirmed transactions of the best chain, including the coinbase of each block.

| Column | Type | Rules |
|---|---|---|
| TxID | Text | Key |
| BlockHeight | Int | Must exist in tblBlocks |
| Position | Int | 0 = coinbase, 1.. = order inside the block. Merkle order |
| FromAddress | Text | Address, or `COINBASE` |
| FromPubKey | Text | Empty for the coinbase |
| ToAddress | Text | Address |
| Amount | Int | Coinbase: reward + fees of the block |
| Fee | Int | 0 for the coinbase |
| Nonce | Int | Coinbase: the block height |
| Timestamp | Int | Unix seconds. Coinbase: the block timestamp |
| Signature | Text | Empty for the coinbase |

## tblBalances (sheet: Balances)

Rebuilt by replaying the whole best chain after every change. Never edited by hand.

| Column | Type | Rules |
|---|---|---|
| Address | Text | Key. This node's wallets first, then every other address on the chain |
| Name | Text | Wallet name if it is a wallet of this node, else empty (names are never shared between nodes) |
| Balance | Int | Everything received minus everything sent (amount + fee) |
| Spendable | Int | Balance minus coinbase coins that are still immature for the NEXT block (tip + 1) |
| NextNonce | Int | Number of confirmed transactions sent by this address |

## tblSyncLog (sheet: SyncLog)

| Column | Type | Rules |
|---|---|---|
| FileName | Text | File name without path |
| Kind | Enum | `Block`, `Tx` |
| ItemID | Text | Block hash or TxID. Empty if the file could not be read |
| Result | Enum | `Added`, `Duplicate`, `Rejected`, `Orphan`, `ForkStored`, `Reorg` |
| Reason | Text | Why, for Rejected, Orphan and Reorg |
| ProcessedAt | Text | Readable UTC text |

## tblLog (sheet: Log)

| Column | Type | Rules |
|---|---|---|
| Timestamp | Text | Readable UTC text |
| Level | Enum | `Info`, `Warning`, `Error` |
| Procedure | Text | Where, example `svcMining.MineBlock` |
| Message | Text | What happened |

---

## Genesis block (fixed, identical on every node)

| Field | Value |
|---|---|
| Height | 0 |
| Timestamp | `1767225600` |
| PrevHash | 64 times `0` |
| MerkleRoot | `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855` (SHA-256 of the empty string) |
| Nonce | 0 |
| DifficultyBits | 12 |
| Miner | `GENESIS` |
| TxCount | 0 |
| Hash | `bb186cf86b1a4db21af766967d13ae380350e900b6d7b823b16f4a2843426a06` |
| CumulativeWork | 4096 |
| Source | `Genesis` |

Hash input: `0|1767225600|0000000000000000000000000000000000000000000000000000000000000000|e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855|0|12|GENESIS`.
Genesis is exempt from the leading-zero check. If `RunSetup` computes a different hash, the hash function is wrong.

---

## Formulas (exact, do not change)

`SHA256(x)` is the lowercase hex SHA-256 of the ASCII text `x`. Integers are plain decimal with no
separators and no leading zeros. `|` is the literal pipe character.

| Item | Formula |
|---|---|
| Address | `"FC" & Left(SHA256(PublicKey), 40)` (PublicKey as its base64 ASCII text) |
| TxID | `SHA256(FromAddress & "|" & ToAddress & "|" & Amount & "|" & Fee & "|" & Nonce & "|" & Timestamp)` |
| Signature | RSA-1024 PKCS#1 v1.5 SHA-256 signature over the ASCII bytes of the 64 character TxID, base64 |
| Merkle leaf | the TxID of each transaction, coinbase first |
| Merkle root | No transactions: `SHA256("")`. One transaction: that TxID. Otherwise pair neighbours with `SHA256(left & right)`; an odd last item is paired with itself; repeat until one hash is left |
| Block hash | `SHA256(Height & "|" & Timestamp & "|" & PrevHash & "|" & MerkleRoot & "|" & Nonce & "|" & DifficultyBits & "|" & Miner)` |
| Halving | `reward(height) = BlockReward \ 2 ^ ((height - 1) \ HalvingInterval)` (integer division) |
| Work of a block | `2 ^ DifficultyBits` |
| Leading zero bits | each leading `0` hex digit counts 4; the first non-zero digit adds `4 - (bits needed to write it)` |

Examples: leading zero bits of `0fff` = 4, `0123` = 7, `1000` = 3, `8000` = 0, `00a0` = 8, `0000` = 16.

Coinbase: `FromAddress = "COINBASE"`, `FromPubKey` and `Signature` empty, `ToAddress` = miner address,
`Amount` = reward + fees, `Fee = 0`, `Nonce` = block height, `Timestamp` = block timestamp.

Required difficulty bits for a block at height H: H = 0 fixed at 12; if `H mod DifficultyInterval <> 0`
the parent's bits; otherwise adjust from the parent's bits `b` using
`expected = TargetBlockSeconds * gaps` and `actual = Timestamp(H-1) - Timestamp(start)`
(start = H - DifficultyInterval, or 1; gaps = H - 1 - start; actual at least 1):
`actual*4 <= expected` -> b+2; `actual*2 <= expected` -> b+1; `actual >= expected*4` -> b-2;
`actual >= expected*2` -> b-1; else b; then clamp to `MinDifficultyBits .. MaxDifficultyBits`.
Blocks that come faster than the target make mining harder; slower blocks make it easier.

---

## Shared folder file formats

Plain ASCII text, CRLF, one item per line. Files are never changed after they are written.
A file with a wrong header is `Rejected`.

### Transaction file: `tx_<first 16 chars of TxID>.txt`

```
FAKECOIN-TX
TX|TxID|FromAddress|FromPubKey|ToAddress|Amount|Fee|Nonce|Timestamp|Signature
```

### Block file: `block_<height as 6 digits>_<first 12 chars of Hash>.txt`

```
FAKECOIN-BLOCK
Height=3
Timestamp=1767226000
PrevHash=<64 hex>
MerkleRoot=<64 hex>
Nonce=4821
DifficultyBits=12
Miner=<address>
Hash=<64 hex>
TxCount=3
TX|TxID|FromAddress|FromPubKey|ToAddress|Amount|Fee|Nonce|Timestamp|Signature
TX|...
```

One `TX|...` line per transaction, coinbase first. For the coinbase line `FromPubKey` and `Signature`
are empty (the separators stay). `CumulativeWork` is not in the file. A reader rejects a file whose
header is wrong, whose fields are missing, or whose TxCount does not match the number of TX lines.
Files are written under a `tmp_` name and renamed, so other nodes never see half a file.

---

## Sample data

None. Each node creates its own wallets with `RunCreateWallet`. A new workbook starts with the
genesis block only; the Mempool, Transactions, Balances and SyncLog tables are empty.
