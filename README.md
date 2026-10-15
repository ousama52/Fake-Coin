# FakeCoin (toy cryptocurrency in Excel VBA)

**FOR LEARNING ONLY.** Key files are not encrypted, RSA-1024 is weak, there is no real network and no
mining competition. Never use this for anything of value.

This README is the full specification: what the program does, how to install and use it, the rules every
node must follow, and the rules for changing the code. The table layouts, exact formulas and file formats
are in `DATA_TABLES.md`; follow them byte for byte, because several workbooks (nodes) must agree exactly.

## Contents

1. What it is
2. Install
3. Using the Control sheet
4. Several nodes
5. Signing and addresses
6. Consensus rules
7. Most-work rule and forks
8. Mining
9. Mempool
10. Shared folder (the fake network)
11. Architecture and module list
12. Rules for the code
13. Error handling pattern
14. Lessons learned
15. Manual checklist
16. Still not like a real cryptocurrency

---

## 1. What it is

A learning project: a toy cryptocurrency that runs in Excel VBA. It works like a real cryptocurrency in
these ways: real RSA signatures (only the owner of a private key can spend), addresses derived from public
keys, difficulty measured in bits and adjusted from block times, the chain with the most cumulative work
wins, coinbase coins mature before they can be spent, a limited Mempool, and confirmations.

Several workbooks (nodes) share one folder as the network. Data lives in Excel Tables. Deliverable:
plain-text `.bas` and `.cls` files in `src/` that are imported into an `.xlsm` workbook.

## 2. Install

1. Open `FakeCoin_Starter.xlsx` and save it as `.xlsm` (Excel Macro-Enabled Workbook). A blank workbook works too.
2. In the VBA editor (Alt+F11) use File > Import File and import the files in `src/` in this order:
   1. Classes: `clsResult`, `clsLogger`, `clsErrorHandler`, `clsBlock`, `clsTransaction`
   2. `modConstants`, `modMessages`, `modSheetUtils`, `modPerformance`, `modTime`
   3. `svcHash`, `svcCrypto`, `svcMerkle`, then the repositories `repoConfig`, `repoWallet`, `repoMempool`,
      `repoBlock`, `repoTransaction`, `repoBalance`, `repoSyncLog`, then the other services `svcWallet`,
      `svcDifficulty`, `svcBalances`, `svcTransaction`, `svcValidation`, `svcMining`, `svcNetwork`, `svcConsensus`
   4. Last: `modSetup`, `modMain`
3. Debug > Compile VBAProject (no errors expected).
4. Click **Setup** on the Control sheet (or run `modMain.RunSetup`). It creates only what is missing and never
   clears data, so it is safe to run again. On the starter workbook everything is already in place; on a
   blank workbook Setup creates the sheets, tables, config keys, genesis block and Control sheet.
5. Create the folder named in `Config.SharedFolder` (default `C:\FakeCoin\network`). Wallet folders are created
   automatically under `Config.WalletFolder` (default `C:\FakeCoin\wallets`), one sub-folder per `Config.NodeID`.

### Hashing and signing on a PC without .NET 3.5

`svcHash.Sha256` tries .NET COM and otherwise uses a built-in pure-VBA SHA-256 (identical output).
`svcCrypto` uses the Windows CryptoAPI for RSA-1024 / SHA-256 signatures (the .NET COM RSA class needs
.NET 3.5 as well). `RunSetup` logs `hash engine: COM` or `VBA`.

## 3. Using the Control sheet

| Button / macro | Inputs (Control cells) | What it does |
|---|---|---|
| Setup / `RunSetup` | | Creates anything missing; never clears data |
| Create wallet / `RunCreateWallet` | C10 name | Real RSA key pair, key file on this node only, wallet row |
| Send coins / `RunSendCoins` | C4 from (wallet of this node), C5 to (name or address), C6 amount, C7 fee, C13 nonce override (TEST ONLY) | Signs, checks and adds to the Mempool, writes a `tx_` file |
| Mine block / `RunMineBlock` | C8 miner (wallet name or any address) | Mines one block at the required difficulty (Esc stops) |
| Sync / `RunSync` | | Reads new files, applies the most-work rule |
| Validate chain / `RunValidateChain` | | Checks every block from genesis (signatures, timestamps, difficulty), reports the first bad height |
| Rebuild from folder / `RunRebuildFromFolder` | | Resets to genesis and rebuilds from the files |
| Tamper test / `RunTamperTest` | C9 height | Changes one transaction amount (local tables only) |
| Tx status / `RunTxStatus` | C12 TxID | `Pending`, `Confirmed, N confirmations`, or `Unknown` |
| Forge test / `RunForgeTest` | C4 victim, C5 receiver, C6 amount, C7 fee | ATTACK SIMULATION: a transaction signed with the wrong key |
| Reset local chain / `RunResetLocal` | | Genesis only (asks first); wallets and key files are kept |
| Reset network folder / `RunResetNetwork` | | Deletes all `block_` and `tx_` files (asks first) |

C11 shows the last run status. C13 (nonce override) is TEST ONLY: it writes a transaction file with a
chosen nonce to the shared folder without adding it to this node's Mempool, so you can create two
conflicting transactions for the double-spend test; the status line warns when it is used.

## 4. Several nodes

1. Run Setup once, save copies: `NodeA.xlsm`, `NodeB.xlsm`, `NodeC.xlsm`. Set `Config.NodeID` differently in each.
2. Keep `SharedFolder` and every NETWORK value (the keys marked NETWORK in `DATA_TABLES.md`) identical in all copies.
3. Each node creates its own wallets: alice on A, bob on B, carol on C. Key files exist only in the node's own folder.
4. Send to another node's **address** (Balances sheet, or the wallet creation message). Click Sync on the other nodes to see the files.
5. For the first tests set `CoinbaseMaturity` to 0 on every node.
6. Balances shows a Name only for this node's own wallets; other addresses appear without a name (names are never shared).

## 5. Signing and addresses

- Keys: RSA-1024 from the Windows CryptoAPI. Public key = base64 of the 128 byte modulus (exponent must be 65537). Address = `FC` + first 40 hex of `SHA256(PublicKey)`.
- A transaction is signed over its TxID text with the sender's private key (read from the key file) and verified with the public key carried in the transaction. The address must match the public key.
- The private key lives only in `<WalletFolder>\<NodeID>\<name>.key` (line 1 `FAKECOIN-KEY`, line 2 the private key as base64). Never in the workbook, never in the shared folder. Existing key files are never overwritten.
- Wallet names use letters, digits and `_` only and are unique on a node.
- Mining is separate from wallets: the miner cell accepts a wallet name or ANY address starting with `FC`.
- To send, the sender must be one of this node's wallets (its key file must exist), otherwise: "No private key for this wallet on this node."

## 6. Consensus rules (every node must apply these identically)

Network-wide constants live in Config and must be the same on every node: `BlockReward`, `HalvingInterval`,
`MaxTxPerBlock`, `MinFee`, `InitialDifficultyBits`, `TargetBlockSeconds`, `DifficultyInterval`,
`MinDifficultyBits`, `MaxDifficultyBits`, `CoinbaseMaturity`, `MaxFutureSeconds`. Genesis is fixed (see
`DATA_TABLES.md`) and identical on every node.

A **block** is valid only if:
1. height = parent height + 1 and `PrevHash` equals the parent's hash;
2. `DifficultyBits` equals the value computed from the chain (never taken from Config, except the fixed genesis value);
3. the hash recomputes and has at least `DifficultyBits` leading zero bits (genesis is exempt);
4. the timestamp is later than the parent's and not later than the current UTC time + `MaxFutureSeconds`;
5. `MerkleRoot` and `TxCount` match;
6. exactly one coinbase, first, with no public key and no signature, Nonce = block height, Fee 0, amount = reward + fees of the other transactions, and a TxID that recomputes;
7. at most `MaxTxPerBlock` other transactions;
8. every other transaction is valid, in order.

`reward(height) = BlockReward \ 2 ^ ((height - 1) \ HalvingInterval)` (integer division). Genesis has no reward.

A **transaction** is valid only if:
1. the address equals `FC` + 40 hex of the SHA-256 of the public key (`address does not match public key` otherwise);
2. the signature verifies (`bad signature` otherwise);
3. the TxID recomputes, Amount > 0, Fee >= MinFee, sender <> receiver, the receiver is `FC` + 40 hex;
4. the nonce equals the sender's next nonce at that point of the chain (pending transactions count in the Mempool);
5. the sender's SPENDABLE balance covers Amount + Fee;
6. the TxID is not already on the chain.

Coinbase coins of block H are spendable in a block at height B only if `B - H >= CoinbaseMaturity`
(`CoinbaseMaturity = 0` turns this off). Balances shows both Balance and Spendable.

**Difficulty** is in bits and adjusted every `DifficultyInterval` blocks from the block times
(exact integer formula in `DATA_TABLES.md`): blocks faster than `TargetBlockSeconds` make mining harder,
slower blocks make it easier. **Timestamps** are Unix seconds (UTC). A new block's timestamp is the
current UTC time, but always later than its parent's.

## 7. Most-work rule and forks

- Work of a block = `2 ^ DifficultyBits`. Every node computes cumulative work itself and never trusts a work value from a table or file.
- The chain with the greater cumulative work wins. On a tie, the node keeps its own chain.
- Received blocks whose parent is unknown are logged `Orphan` in SyncLog and retried on every Sync.
- A received valid block that does not extend my tip is kept as a fork candidate (`ForkStored`, the files stay in the folder; there is no side table). The candidate branch is built by following `PrevHash` through the block files back to a block on my chain (the fork point).
- If the candidate branch has MORE cumulative work than my chain: replay the whole new chain from genesis, validating every block. If everything is valid, switch to it; otherwise keep my chain and log `Rejected` with the reason.
- On a switch, transactions from my dropped blocks that are not in the new chain go back to the Mempool (if still valid), and Mempool entries now confirmed are removed. Transactions and Balances are rebuilt from the new chain. A `Reorg` line is logged with the old tip, the new tip and the fork height.

## 8. Mining

- Take the Mempool, keep only transactions that pass the transaction rules in order (signature, nonce, spendable balance), highest fee first, up to `MaxTxPerBlock`. After each pick the scan restarts, so a sender's nonce 1 can follow its nonce 0 even when it pays the higher fee.
- Build the coinbase for the miner address (reward + fees), compute the Merkle root, then loop `Nonce` from 0 until the hash has the required leading zero bits. `DoEvents` every 5000 tries so Excel stays responsive; Esc stops.
- Build and check the block in memory, then write once: Blocks, Transactions, the Mempool, Balances and the `block_` file. An empty Mempool still mines a block with only the coinbase.

## 9. Mempool

- At most `MempoolMaxTx` entries (local setting). When full, a transaction with a higher fee than the lowest-fee entry replaces that entry; otherwise it is rejected with "Mempool full, fee too low." (The replaced sender's later nonces can become stuck.)
- A transaction is admitted only after all the transaction rules pass, using pending nonces and spendable balance. A transaction file from the network is checked the same way during Sync.
- A transaction whose nonce is ahead of the sender's pending count is logged `Orphan` and retried on every Sync.

## 10. Shared folder (the fake network)

- Path in `Config.SharedFolder`. All nodes point to the same folder; `Config.NodeID` names this node. File names and line formats are in `DATA_TABLES.md`. Files are never overwritten: if the name exists, skip it.
- Sync processes block files first, then transaction files, each in name order. It skips names already in SyncLog (except `Orphan`). It never stops on a bad file: log it as `Rejected` and carry on (wrong header, missing fields, garbled content, TxCount mismatch).
- Files are written under a `tmp_` name and renamed, so other nodes never read half a file.
- Edge cases handled: the same block file seen twice, or a block already on my chain, is `Duplicate`; two nodes mining at the same height while disconnected end on the same chain after Sync; re-running Setup never destroys data; empty tables never crash array reads (an empty table keeps one blank row, and a row whose first column is empty counts as no data).

## 11. Architecture and module list

```
Entry (modMain)
  -> Services (svc*)            rules: hashing, crypto, wallets, transactions, mining, difficulty, validation, consensus, network
      -> Repositories (repo*)   read/write tables in bulk
          -> Infrastructure (clsLogger, clsErrorHandler, modSheetUtils, modPerformance, modTime)
Models (cls*) are plain data holders.
```

- Entry macros only orchestrate: error handling, call a service, show a message. No rules.
- Services never touch worksheets directly. They use repositories. `svcNetwork` is the only module that touches the shared folder. `svcWallet` writes and reads key files.
- Nothing below Entry shows a MsgBox. Services return a `clsResult` (Success, Message, Count).

| Layer | Modules |
|---|---|
| Entry | `modMain` |
| Config | `modConstants`, `modMessages`, `modSetup` |
| Services | `svcHash`, `svcCrypto`, `svcWallet`, `svcTransaction`, `svcMerkle`, `svcDifficulty`, `svcMining`, `svcValidation`, `svcConsensus`, `svcNetwork`, `svcBalances` |
| Repositories | `repoConfig`, `repoWallet`, `repoMempool`, `repoBlock`, `repoTransaction`, `repoBalance`, `repoSyncLog` |
| Models | `clsBlock`, `clsTransaction`, `clsResult` |
| Infrastructure | `clsLogger`, `clsErrorHandler`, `modSheetUtils`, `modPerformance`, `modTime` |

## 12. Rules for the code

1. Output ONLY `.bas` and `.cls` files plus `README.md` and `DATA_TABLES.md`. No UserForms, no binary files. The UI is a `Control` worksheet with buttons that call entry macros.
2. Code must be ASCII only (code, comments and files written to the shared folder).
3. Save files with CRLF line endings.
4. Every `.bas` file starts with `Attribute VB_Name = "modName"` then `Option Explicit`.
5. Every `.cls` file starts with the full class header (`VERSION 1.0 CLASS`, `BEGIN`, `MultiUse = -1  'True`, `END`, `Attribute VB_Name`, `VB_GlobalNameSpace = False`, `VB_Creatable = False`, `VB_PredeclaredId = False`, `VB_Exposed = False`), then `Option Explicit`. `clsLogger` and `clsErrorHandler` use `VB_PredeclaredId = True` so entry macros can call them without an object.
6. All public names are unique across the project and prefixed by layer (`svc`, `repo`, `cls`, `mod`). Always call with the module name.
7. Coin amounts are whole numbers (`Long`). No decimals. Cumulative work uses `Decimal` (`CDec`). Never format numbers with separators inside hash input.
8. No automated tests, test modules or test frameworks in the project. Verify with the manual checklist below.
9. No `Select`, `Activate`, `ActiveSheet`, `ActiveCell` or `Selection`. Refer to sheets and tables by name through `modConstants`.
10. No cell-by-cell loops over table data. Read a table into a 2D array, work in memory, write back in one assignment.
11. Windows desktop Excel only. `svcHash.Sha256(text)` is the ONLY SHA-256 function; nothing else may hash. `svcCrypto` is the ONLY module that does RSA. Base64 goes through `MSXML2.DOMDocument`.
12. A VBA hex literal of 1 to 4 digits is a signed Integer: write `&H800C&` when a Long is meant.
13. Ask before changing any hash formula, file format or table column.
14. Keep procedures short (about 60 lines max), one responsibility each. Comment the why, not the what.
15. Keep the "FOR LEARNING ONLY" notices on the Control sheet, the Wallets sheet, in this README and at the top of `svcWallet` and `svcCrypto`. Do not optimize or harden beyond what is written here.

## 13. Error handling pattern

Every entry macro uses this shape. The error text is copied first because logging can overwrite `Err`.

```vba
Public Sub RunMineBlock()
    Dim result As clsResult
    Dim errText As String
    On Error GoTo ErrHandler
    modPerformance.TurnOff
    Set result = svcMining.MineBlock(repoConfig.GetMinerFromControl())
    ShowResult result
Cleanup:
    On Error Resume Next
    modPerformance.TurnOn          ' restores ScreenUpdating, Calculation, EnableEvents, status bar
    Exit Sub
ErrHandler:
    errText = Err.Description
    clsErrorHandler.LogError "RunMineBlock", Err
    MsgBox "Error: " & errText, vbCritical
    Resume Cleanup
End Sub
```

Build changes in memory first, then write once: an error halfway must never leave a half-written block.
Log every major step to the Log sheet through `clsLogger`.

## 14. Lessons learned

- Class public fields cannot be passed `ByRef` (VBA passes a copy): parse into a local, then assign.
- An Excel Table with no data rows can have `DataBodyRange = Nothing`: write by position below the header and resize afterwards; put validation on the column below the header.
- `NumberFormat` minutes are `mm` (not `nn`); `line` is a reserved word; `ReDim Preserve` cannot resize the first dimension of a 2D array.
- If `CreateObject` for a .NET class fails with error 80131700, the old .NET 3.5 runtime is missing; that is why hashing falls back to VBA and RSA uses the CryptoAPI.

## 15. Manual checklist

1. Compile has no errors. `?svcHash.Sha256("abc")` gives `ba7816bf8f01cfea414140de5dae2223b00361a396177a9cb410ff61f20015ad`. `?svcHash.LeadingZeroBits("0123")` gives 7.
2. Sign and verify: create a key pair with `svcCrypto.GenerateKeyPair`, sign a string, `VerifySignature` is True; change one character and it is False.
3. Setup leaves one genesis block with hash `bb186cf86b1a4db21af766967d13ae380350e900b6d7b823b16f4a2843426a06` and CumulativeWork 4096.
4. Set `CoinbaseMaturity` to 0 for the early tests. Node A creates `alice`, B creates `bob`, C creates `carol`. Key files exist only in their own node folders.
5. Mine on A paying alice (50). Sync B and C. Same chain on all nodes.
6. alice sends to bob's address with fee 1. Sync B, mine on B paying carol's address. Balances match on all nodes. Tx status shows the right confirmations after more blocks.
7. Forgery: run Forge test on B using alice's address. After Sync, A rejects it as `bad signature` or `address does not match public key`. Nobody's balance changes.
8. Double spend: alice sends 40 to bob on A. Use the nonce override to make a second alice transaction with the same nonce but a different receiver. Keep the first transaction file away from B, Sync B, mine on A and B separately, then mine a second block on A and Sync everywhere. Exactly one of the two payments survives and every node agrees.
9. Cheating: Tamper test, then Validate chain finds it, and Rebuild from folder repairs it.
10. Difficulty: set `DifficultyInterval` 4 and `TargetBlockSeconds` very high (for example 3600) on every node before a fresh chain. Mine quickly: at height 4 the required bits go UP by 2. With a very low target (for example 1) and slow mining the bits go DOWN.
11. Maturity: set `CoinbaseMaturity` to 3. alice mines a block and cannot spend until blocks are on top of it.
12. Fork: disconnect A and B, mine one block on each, then a second on A. B switches to A's chain (more cumulative work). The `Reorg` line appears in SyncLog.
13. Bad timestamps: edit a copy of a block file so the timestamp is before the parent. The other node rejects it. A node whose clock is an hour ahead mines a block: the others reject it as too far in the future.
14. Wrong-header and garbled files in the folder are `Rejected` and Sync carries on.
15. No leftover `EXCEL.EXE` in Task Manager.

## 16. Still not like a real cryptocurrency

- RSA instead of ECDSA, 1024-bit keys, unencrypted key files, private data held in VBA memory
- No real peer-to-peer networking: one shared folder, files read on Sync
- You click Mine, no competition between miners
- A handful of nodes instead of thousands
- No UTXO model, no scripts, no light clients, no mempool relay rules beyond the ones above
