var Constants = {
    Column: {
        DATE: 1,
        TYPE: 2,
        AMOUNT: 3,
        BALANCE: 4,
        MEMO: 5,
    },

    HistoryRow: {
        FIRST_DATA_ROW: 2
    },

    InterestRateLocation: {
        column: 3,
        row: 2,
    },

    LedgerRow: {
        FIRST_DATA_ROW: 5
    },

    NUM_DATA_COLUMNS: 5,
    NUM_REQ_DATA_COLUMNS: 4,

    SheetName: {
        HISTORY: 'History',
        LEDGER: 'Ledger',
    }
};

function getConstants() {
    return Constants;
}
