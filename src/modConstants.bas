Attribute VB_Name = "modConstants"
Option Explicit

' FAKE - FOR LEARNING ONLY. Nothing in this project is secure.

' Worksheet names (sheet = table name without the tbl prefix)
Public Const SHEET_CONFIG As String = "Config"
Public Const SHEET_WALLETS As String = "Wallets"
Public Const SHEET_MEMPOOL As String = "Mempool"
Public Const SHEET_BLOCKS As String = "Blocks"
Public Const SHEET_TRANSACTIONS As String = "Transactions"
Public Const SHEET_BALANCES As String = "Balances"
Public Const SHEET_SYNCLOG As String = "SyncLog"
Public Const SHEET_LOG As String = "Log"
Public Const SHEET_CONTROL As String = "Control"

' Table names
Public Const TBL_CONFIG As String = "tblConfig"
Public Const TBL_WALLETS As String = "tblWallets"
Public Const TBL_MEMPOOL As String = "tblMempool"
Public Const TBL_BLOCKS As String = "tblBlocks"
Public Const TBL_TRANSACTIONS As String = "tblTransactions"
Public Const TBL_BALANCES As String = "tblBalances"
Public Const TBL_SYNCLOG As String = "tblSyncLog"
Public Const TBL_LOG As String = "tblLog"

' Config keys
Public Const CFG_COIN_NAME As String = "CoinName"
Public Const CFG_NODE_ID As String = "NodeID"
Public Const CFG_SHARED_FOLDER As String = "SharedFolder"
Public Const CFG_WALLET_FOLDER As String = "WalletFolder"
Public Const CFG_BLOCK_REWARD As String = "BlockReward"
Public Const CFG_HALVING As String = "HalvingInterval"
Public Const CFG_MAX_TX As String = "MaxTxPerBlock"
Public Const CFG_MIN_FEE As String = "MinFee"
Public Const CFG_INIT_BITS As String = "InitialDifficultyBits"
Public Const CFG_TARGET_SECONDS As String = "TargetBlockSeconds"
Public Const CFG_DIFF_INTERVAL As String = "DifficultyInterval"
Public Const CFG_MIN_BITS As String = "MinDifficultyBits"
Public Const CFG_MAX_BITS As String = "MaxDifficultyBits"
Public Const CFG_MATURITY As String = "CoinbaseMaturity"
Public Const CFG_MAX_FUTURE As String = "MaxFutureSeconds"
Public Const CFG_MEMPOOL_MAX As String = "MempoolMaxTx"

' Positions in the network-parameter array returned by repoConfig.LoadChainParams
Public Const PARAM_REWARD As Long = 1
Public Const PARAM_HALVING As Long = 2
Public Const PARAM_MAXTX As Long = 3
Public Const PARAM_MINFEE As Long = 4
Public Const PARAM_INIT_BITS As Long = 5
Public Const PARAM_TARGET_SECONDS As Long = 6
Public Const PARAM_DIFF_INTERVAL As Long = 7
Public Const PARAM_MIN_BITS As Long = 8
Public Const PARAM_MAX_BITS As Long = 9
Public Const PARAM_MATURITY As Long = 10
Public Const PARAM_MAX_FUTURE As Long = 11
Public Const PARAM_MEMPOOL_MAX As Long = 12
Public Const PARAM_COUNT As Long = 12

' Enum values
Public Const SRC_LOCAL As String = "Local"
Public Const SRC_NETWORK As String = "Network"
Public Const SRC_GENESIS As String = "Genesis"
Public Const SRC_MINED As String = "Mined"

Public Const SYNC_KIND_BLOCK As String = "Block"
Public Const SYNC_KIND_TX As String = "Tx"

Public Const SYNC_ADDED As String = "Added"
Public Const SYNC_DUPLICATE As String = "Duplicate"
Public Const SYNC_REJECTED As String = "Rejected"
Public Const SYNC_ORPHAN As String = "Orphan"
Public Const SYNC_FORK_STORED As String = "ForkStored"
Public Const SYNC_REORG As String = "Reorg"

Public Const LOG_INFO As String = "Info"
Public Const LOG_WARNING As String = "Warning"
Public Const LOG_ERROR As String = "Error"

' Result codes of svcTransaction.AdmitToMempool
Public Const ADM_OK As Long = 0
Public Const ADM_DUP As Long = 1
Public Const ADM_REJECT As Long = 2
Public Const ADM_WAIT As Long = 3

' Fixed strings of the protocol
Public Const COINBASE_KEY As String = "COINBASE"
Public Const ADDRESS_PREFIX As String = "FC"
Public Const ADDRESS_LENGTH As Long = 42
Public Const FILE_BLOCK_HEADER As String = "FAKECOIN-BLOCK"
Public Const FILE_TX_HEADER As String = "FAKECOIN-TX"
Public Const KEY_FILE_HEADER As String = "FAKECOIN-KEY"

' Genesis block (fixed, identical on every node). Timestamp is Unix seconds.
Public Const GENESIS_TIMESTAMP As Long = 1767225600
Public Const GENESIS_PREVHASH As String = "0000000000000000000000000000000000000000000000000000000000000000"
Public Const GENESIS_MERKLE As String = "e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855"
Public Const GENESIS_MINER As String = "GENESIS"
Public Const GENESIS_BITS As Long = 12
Public Const GENESIS_WORK As Long = 4096
Public Const GENESIS_HASH As String = "bb186cf86b1a4db21af766967d13ae380350e900b6d7b823b16f4a2843426a06"

' Control sheet cells
Public Const CTL_FROM As String = "C4"
Public Const CTL_TO As String = "C5"
Public Const CTL_AMOUNT As String = "C6"
Public Const CTL_FEE As String = "C7"
Public Const CTL_MINER As String = "C8"
Public Const CTL_HEIGHT As String = "C9"
Public Const CTL_WALLET As String = "C10"
Public Const CTL_STATUS As String = "C11"
Public Const CTL_TXID As String = "C12"
Public Const CTL_NONCE As String = "C13"

' Base for application error numbers raised with Err.Raise
Public Const ERR_APP As Long = vbObjectError + 2000
