function fakelog(text) {
    const range = getLedger().getRange(1,10);
    range.setValues([[range.getValue()+'\n'+text]]);
}

function archiveHistory() {
    const ledger = getLedger();
    if(isEmpty(ledger)) {
        return;
    }

    const history = getHistory();
    // Validate that first row in ledger is correct according to history
    const historicBalance = getHistoricBalance();

    const firstChange = ledger.getRange(Constants.LedgerRow.FIRST_DATA_ROW, Constants.Column.AMOUNT).getValue();
    if (historicBalance + firstChange !== getStartingBalance(ledger)) {
        // throw new Error(`Historic balance is incorrect: ${historicBalance} + ${firstChange} !== ${getStartingBalance()}`);
        throw new Error('Historic balance is incorrect');
    }

    // Find the last row to archive
    const startOfCurrentMonth = getStartOfMonth(new Date());
    const range = ledger.getDataRange().getValues();
    var lastRow = null;
    var fullArchive = false;
    const firstRow = getFirstDataRow(ledger);
    for (var row = range.length - 1; row > firstRow && lastRow === null; row--) {
        if (isBefore(convertDateToMidnight(range[row][0]), startOfCurrentMonth)) {
            if (row === range.length - 1) {
                fullArchive = true;
            }
            lastRow = row+1;
        }
    }

    // If there's no data earlier than this month, do nothing
    if (lastRow === null) {
        return;
    }

    const numRows = lastRow - firstRow + 1;
    const lastHistoricRow = history.getLastRow();
    const currentBalance = getCurrentBalance();
    const firstLineIsStartingBalance = ledger.getRange(firstRow,2).getValue() === 'Starting Balance';

    var srcRange = ledger.getRange(firstLineIsStartingBalance ? firstRow+1 : firstRow, 1, numRows, Constants.NUM_DATA_COLUMNS);
    var destRange = history.getRange(lastHistoricRow+1, 1, numRows, 5);
    srcRange.copyTo(destRange);
    srcRange.clearContent();
    if (fullArchive) {
        // Add a starting balance row
        ledger.getRange(firstRow, 1, 1, 4).setValues([[startOfCurrentMonth, 'Starting Balance', 0, currentBalance]]);
    } else {
        // Move all the remaining rows up
        const oldLastRow = ledger.getLastRow();
        const remainingRows = oldLastRow - lastRow;
        const remainingLedgerRange = ledger.getRange(lastRow + 1, 1, remainingRows, Constants.NUM_DATA_COLUMNS);
        remainingLedgerRange.copyTo(ledger.getRange(firstRow, 1, remainingRows, Constants.NUM_DATA_COLUMNS));
        remainingLedgerRange.clearContent();
    }
}

function fullProcessing() {
    recalculateBalances();
    generateInterestPayments();
    archiveHistory();
}