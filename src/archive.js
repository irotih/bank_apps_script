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

    const firstChange = ledger.getRange(2,3).getValue();
    if (historicBalance + firstChange !== getStartingBalance(ledger)) {
        // throw new Error(`Historic balance is incorrect: ${historicBalance} + ${firstChange} !== ${getStartingBalance()}`);
        throw new Error('Historic balance is incorrect');
    }

    // First calculate all missing interest payments
    // This will also ensure that balances are reported up to the last month
    generateInterestPayments();

    // Find the last row to archive
    const startOfCurrentMonth = getStartOfMonth(new Date());
    fakelog('startOfCurrentMonth = '+startOfCurrentMonth);
    const range = ledger.getDataRange().getValues();
    var lastRow = null;
    var fullArchive = false;
    for (var row = range.length - 1; row > 0 && lastRow === null; row--) {
        fakelog('row '+row+', date: '+range[row][0]);
        if (isBefore(convertDateToMidnight(range[row][0]), startOfCurrentMonth)) {
            if (row === range.length - 1) {
                fullArchive = true;
            }
            lastRow = row+1;
        }
    }

    fakelog('Last row to archive: '+lastRow);

    // If there's no data earlier than this month, do nothing
    if (lastRow === null) {
        return;
    }

    const numRows = lastRow - 1;
    const lastHistoricRow = history.getLastRow();
    fakelog('last historic row = '+lastHistoricRow);
    const currentBalance = getCurrentBalance();
    const firstLineIsStartingBalance = ledger.getRange(2,2).getValue() === 'Starting Balance';

    var srcRange = ledger.getRange(firstLineIsStartingBalance ? 3 : 2, 1, numRows, 5);
    var destRange = history.getRange(lastHistoricRow+1, 1, numRows, 5);
    srcRange.copyTo(destRange);
    srcRange.clearContent();
    if (fullArchive) {
        // Add a starting balance row
        ledger.getRange(2, 1, 1, 4).setValues([[startOfCurrentMonth, 'Starting Balance', 0, currentBalance]]);
    } else {
        // Move all the remaining rows up
        const oldLastRow = ledger.getLastRow();
        fakelog('oldLastRow = '+oldLastRow);
        const remainingLedgerRange = ledger.getRange(lastRow + 1, 1, oldLastRow, 5);
        for (var row = 1; lastRow + row <= oldLastRow; row++) {
            ledger.getRange(lastRow + row, 1, 1, 5).copyTo(ledger.getRange(row+1, 1, 1, 5));
        }
        remainingLedgerRange.clearContent();
    }
}
