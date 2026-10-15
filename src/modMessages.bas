Attribute VB_Name = "modMessages"
Option Explicit

' User-facing text. Placeholders are {0}, {1}, ... filled by FormatText.
' FOR LEARNING ONLY.

Public Const MSG_SHEET_MISSING As String = "Sheet '{0}' was not found. Run Setup first."
Public Const MSG_TABLE_MISSING As String = "Table '{0}' was not found. Run Setup first."
Public Const MSG_COLUMN_MISMATCH As String = "Data has {0} column(s) but table '{1}' has {2}."
Public Const MSG_ERROR As String = "Error: {0}"
Public Const MSG_CANCELLED As String = "Cancelled."

' Setup
Public Const MSG_SETUP_DONE As String = "Setup complete. Items created: {0}. Existing data was not changed."
Public Const MSG_GENESIS_BAD As String = "Genesis hash mismatch: computed {0} but expected {1}. The hash function is wrong."

' Cryptography and wallets
Public Const MSG_CRYPTO_FAIL As String = "Windows cryptography call '{0}' failed (error {1})."
Public Const MSG_CRYPTO_EXPONENT As String = "The public exponent is not 65537."
Public Const MSG_CRYPTO_BAD_KEY As String = "The key text is not valid."
Public Const MSG_WALLET_CREATED As String = "Wallet {0} created. Address {1}. Private key file: {2} (FOR LEARNING ONLY: the key file is not encrypted)."
Public Const MSG_WALLET_NAME_EMPTY As String = "Type a wallet name in Control!C10 first."
Public Const MSG_WALLET_BADNAME As String = "A wallet name can only use letters, digits and _."
Public Const MSG_WALLET_EXISTS As String = "A wallet named '{0}' already exists on this node."
Public Const MSG_WALLET_KEYFILE_EXISTS As String = "A key file already exists and is never overwritten: {0}"
Public Const MSG_WALLET_UNKNOWN As String = "'{0}' is not a wallet name or an address (FC + 40 hex characters)."
Public Const MSG_WALLET_NOKEY As String = "No private key for this wallet on this node."

' Transactions
Public Const MSG_TX_SENT As String = "Transaction {0} added to the Mempool: {1} sends {2} to {3}, fee {4}, nonce {5}."
Public Const MSG_TX_FILE_ONLY As String = "TEST ONLY: nonce override {0} used. Transaction {1} was written to the shared folder but NOT added to this node's Mempool."
Public Const MSG_TX_R_FROM As String = "Sender wallet '{0}' is not in the Wallets sheet of this node."
Public Const MSG_TX_R_AMOUNT As String = "Amount must be a whole number greater than 0."
Public Const MSG_TX_R_FEE As String = "Fee must be a whole number of at least {0}."
Public Const MSG_TX_R_NONCE_INPUT As String = "The nonce override must be a whole number."
Public Const MSG_TX_R_SELF As String = "Sender and receiver are the same."
Public Const MSG_TX_R_TOADDR As String = "ToAddress must be FC followed by 40 hex characters."
Public Const MSG_TX_R_BALANCE As String = "Not enough spendable balance: available {0}, needed {1}."
Public Const MSG_TX_R_NONCE As String = "Nonce {0} is not the expected {1}."
Public Const MSG_TX_R_ID As String = "TxID does not match the transaction data."
Public Const MSG_TX_R_ADDR As String = "address does not match public key"
Public Const MSG_TX_R_SIG As String = "bad signature"
Public Const MSG_TX_R_DUP As String = "Transaction is already in the chain."
Public Const MSG_TX_R_FIELDS As String = "A transaction field is empty."
Public Const MSG_TX_R_MEMPOOL_FULL As String = "Mempool full, fee too low."
Public Const MSG_TXSTATUS_PENDING As String = "Transaction {0}: Pending"
Public Const MSG_TXSTATUS_CONFIRMED As String = "Transaction {0}: Confirmed, {1} confirmations"
Public Const MSG_TXSTATUS_UNKNOWN As String = "Transaction {0}: Unknown"
Public Const MSG_TXSTATUS_NOINPUT As String = "Type a TxID in Control!C12 first."
Public Const MSG_FORGE_DONE As String = "ATTACK SIMULATION: forged transaction {0}, claiming to be from {1}, was written to the shared folder. Other nodes should reject it with: {2}. Check SyncLog after Sync."

' Mining
Public Const MSG_MINED As String = "Mined block {0}: hash {1}, {2} bits, {3} transaction(s) including the coinbase, nonce {4}, miner receives {5}."
Public Const MSG_MINING_STOPPED As String = "Mining stopped (Esc). Nothing was written."
Public Const MSG_MINER_UNKNOWN As String = "Miner '{0}' is not a wallet name or an address. Type one in Control!C8."

' Validation
Public Const MSG_CHAIN_OK As String = "Chain is valid: {0} block(s), tip height {1}, cumulative work {2}."
Public Const MSG_CHAIN_BAD As String = "Chain is INVALID. First bad height: {0}. Reason: {1}"
Public Const MSG_BLK_GENESIS As String = "the genesis block was changed"
Public Const MSG_BLK_BITS As String = "DifficultyBits {0} but the required value is {1}"
Public Const MSG_BLK_HASH As String = "Hash does not match the block data"
Public Const MSG_BLK_POW As String = "Hash has {0} leading zero bits but needs {1}"
Public Const MSG_BLK_HEIGHT As String = "height {0} is not parent height + 1"
Public Const MSG_BLK_PREV As String = "PrevHash does not match the parent Hash"
Public Const MSG_BLK_TS_PARENT As String = "Timestamp is not later than the parent's"
Public Const MSG_BLK_TS_FUTURE As String = "Timestamp is too far in the future"
Public Const MSG_BLK_TXCOUNT As String = "TxCount {0} does not match {1} transaction(s)"
Public Const MSG_BLK_MERKLE As String = "MerkleRoot does not match the transactions"
Public Const MSG_BLK_NOCB As String = "the first transaction is not a coinbase"
Public Const MSG_BLK_MULTICB As String = "more than one coinbase"
Public Const MSG_BLK_CBNONCE As String = "coinbase Nonce must equal the block height"
Public Const MSG_BLK_CBID As String = "coinbase TxID does not match its data"
Public Const MSG_BLK_CBFEE As String = "coinbase Fee must be 0"
Public Const MSG_BLK_CBFIELDS As String = "coinbase must have no public key and no signature"
Public Const MSG_BLK_CBAMOUNT As String = "coinbase Amount is {0} but must be {1}"
Public Const MSG_BLK_MAXTX As String = "{0} transactions exceed MaxTxPerBlock {1}"
Public Const MSG_BLK_TX As String = "transaction {0}: {1}"
Public Const MSG_BLK_EMPTYCHAIN As String = "the chain has no blocks"
Public Const MSG_TAMPER_DONE As String = "Changed Amount of transaction {0} in block {1} from {2} to {3} (local tables only). Run Validate chain to see it caught."
Public Const MSG_TAMPER_BAD_HEIGHT As String = "Block {0} does not exist. Type a height in Control!C9."
Public Const MSG_TAMPER_GENESIS As String = "Block 0 (genesis) has no transactions. Pick a height of 1 or more."

' Shared folder and files
Public Const MSG_FOLDER_MISSING As String = "Shared folder '{0}' does not exist. Create it or change Config.SharedFolder."
Public Const MSG_FILE_UNREADABLE As String = "file could not be read"
Public Const MSG_FILE_HEADER As String = "wrong header line"
Public Const MSG_FILE_FIELD As String = "missing or invalid field {0}"
Public Const MSG_FILE_TXLINE As String = "bad TX line {0}"
Public Const MSG_FILE_TXCOUNT As String = "TxCount {0} does not match {1} TX line(s)"
Public Const MSG_RESET_NET_DONE As String = "Network folder cleared: {0} file(s) deleted."

' Sync
Public Const MSG_SYNC_DONE As String = "Sync: {0} file(s) read - added {1}, duplicate {2}, rejected {3}, orphan {4}, fork stored {5}, reorg {6}. Tip height {7}."
Public Const MSG_SYNC_ORPHAN As String = "parent block {0} is not known yet"
Public Const MSG_SYNC_NONCE_WAIT As String = "waiting for an earlier nonce from the same sender"
Public Const MSG_SYNC_NONCE_USED As String = "nonce {0} is already used (expected {1})"
Public Const MSG_SYNC_LOOP As String = "the parent links form a loop"
Public Const MSG_REORG As String = "old tip {0} (height {1}, work {2}), new tip {3} (height {4}, work {5}), fork height {6}"
Public Const MSG_REBUILD_DONE As String = "Rebuilt from the shared folder. {0}"
Public Const MSG_RESET_LOCAL_DONE As String = "Local chain reset to genesis. Mempool, Transactions, Balances and SyncLog were cleared. Wallets and key files were kept."

' Prompts
Public Const MSG_CONFIRM_RESET_LOCAL As String = "Reset the local chain to genesis? Mempool, Transactions, Balances and SyncLog are cleared. Wallets and key files are kept."
Public Const MSG_CONFIRM_RESET_NET As String = "Delete ALL block_ and tx_ files in the shared folder? This affects every node."

' Control sheet labels and buttons
Public Const LBL_FAKE As String = "FOR LEARNING ONLY. Key files are unencrypted and this is not a secure cryptocurrency."
Public Const LBL_CTL_TITLE As String = "FakeCoin (toy cryptocurrency)"
Public Const LBL_CTL_FROM As String = "Send from (wallet name) / Forge: victim"
Public Const LBL_CTL_TO As String = "Send to (wallet name or address)"
Public Const LBL_CTL_AMOUNT As String = "Amount (whole coins)"
Public Const LBL_CTL_FEE As String = "Fee (whole coins)"
Public Const LBL_CTL_MINER As String = "Miner (wallet name or address)"
Public Const LBL_CTL_HEIGHT As String = "Block height (tamper test)"
Public Const LBL_CTL_WALLET As String = "New wallet name"
Public Const LBL_CTL_STATUS As String = "Last run status"
Public Const LBL_CTL_TXID As String = "TxID (Tx status)"
Public Const LBL_CTL_NONCE As String = "Nonce override (TEST ONLY, leave blank)"

Public Const BTN_SETUP As String = "Setup"
Public Const BTN_CREATE_WALLET As String = "Create wallet"
Public Const BTN_SEND As String = "Send coins"
Public Const BTN_MINE As String = "Mine block"
Public Const BTN_SYNC As String = "Sync"
Public Const BTN_VALIDATE As String = "Validate chain"
Public Const BTN_REBUILD As String = "Rebuild from folder"
Public Const BTN_TAMPER As String = "Tamper test"
Public Const BTN_TXSTATUS As String = "Tx status"
Public Const BTN_FORGE As String = "Forge test"
Public Const BTN_RESET_LOCAL As String = "Reset local chain"
Public Const BTN_RESET_NET As String = "Reset network folder"

Public Function FormatText(ByVal template As String, ParamArray args() As Variant) As String
    Dim i As Long
    Dim s As String

    s = template
    For i = LBound(args) To UBound(args)
        s = Replace(s, "{" & i & "}", CStr(args(i)))
    Next i
    FormatText = s
End Function
