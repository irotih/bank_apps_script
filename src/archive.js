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
    if (!isEmpty(history) && historicBalance + firstChange !== getStartingBalance(ledger)) {
        throw new Error('Historic balance is incorrect: ' + historicBalance + ' + ' + firstChange + ' != ' + getStartingBalance(ledger));
    }
    // Find the last row to archive
    const currentDate = new Date();
    const historyCutoffDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - Constants.MONTHS_TO_RETAIN, 1);
    const firstRow = getFirstDataRow(ledger);
    const lastRow = ledger.getLastRow();

    const range = ledger.getRange(firstRow, 1, lastRow - firstRow + 1, 1).getValues();

    var lastRowToArchive = null;
    var fullArchive = false;
    for (var row = range.length - 1; row >= 0 && lastRowToArchive === null; row--) {
        if (isBefore(convertDateToMidnight(range[row][0]), historyCutoffDate)) {
            if (row === range.length - 1) {
                fullArchive = true;
            }
            lastRowToArchive = row + firstRow;
        }
    }

    // If there's no data earlier than this month, do nothing
    if (lastRowToArchive === null) {
        return;
    }

    const firstLineIsStartingBalance = ledger.getRange(firstRow,2).getValue() === 'Starting Balance';
    const firstRowToArchive = firstLineIsStartingBalance ? firstRow + 1 : firstRow;
    const numRowsToArchive = lastRowToArchive - firstRowToArchive + 1;

    const historyInsertionRow = history.getLastRow() + 1;
    const currentBalance = getCurrentBalance();

    var srcRange = ledger.getRange(firstRowToArchive, 1, numRowsToArchive, Constants.NUM_DATA_COLUMNS);
    var destRange = history.getRange(historyInsertionRow, 1, numRowsToArchive, Constants.NUM_DATA_COLUMNS);
    srcRange.copyTo(destRange);
    srcRange.clearContent();
    if (fullArchive) {
        // Add a starting balance row
        ledger.getRange(firstRow, 1, 1, 4).setValues([[startOfCurrentMonth, 'Starting Balance', 0, currentBalance]]);
    } else {
        // Move all the remaining rows up
        shiftRowsUp(ledger, lastRowToArchive + 1, ledger.getLastRow(), firstRow);
    }
}

function fullProcessing() {
    recalculateBalances();
    generateInterestPayments();
    archiveHistory();
}