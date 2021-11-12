function archiveHistory() {
  if(isEmpty()) {
    return;
  }

  const ledger = getLedger();
  const history = getHistory();
  // Validate that first row in ledger is correct according to history
  const historicBalance = getHistoricBalance();
  const firstChange = ledger.getRange(2,3).getValue();
  if (historicBalance + firstChange !== getStartingBalance()) {
    // throw new Error(`Historic balance is incorrect: ${historicBalance} + ${firstChange} !== ${getStartingBalance()}`);
    throw new Error('Historic balance is incorrect');
  }

  // First calculate all missing interest payments
  // This will also ensure that balances are reported up to the last month
  generateInterestPayments();

  // Find the last row to archive
  const startOfCurrentMonth = moment().startOf('month');
  const range = ledger.getDataRange().getValues();
  var lastRow = null;
  var fullArchive = false;
  for (var row = range.length - 1; row > 0 && lastRow === null; row--) {
    if (convertToMoment(range[row][0]).isBefore(startOfCurrentMonth)) {
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

  const numRows = lastRow - 1;
  const lastHistoricRow = history.getLastRow();
  const currentBalance = getCurrentBalance();
  const firstLineIsStartingBalance = ledger.getRange(2,2).getValue() === 'Starting Balance';

  var srcRange = ledger.getRange(firstLineIsStartingBalance ? 3 : 2, 1, numRows, 5);
  var destRange = history.getRange(lastHistoricRow+1, 1, numRows, 5);
  srcRange.copyTo(destRange);
  srcRange.clearContent();
  if (fullArchive) {
    // Add a starting balance row
    ledger.getRange(2, 1, 1, 4).setValues([[startOfCurrentMonth.toDate(), 'Starting Balance', 0, currentBalance]]);
  } else {
    // Move all the remaining rows up
    const oldLastRow = ledget.getLastRow();
    const remainingLedgerRange = ledger.getRange(lastRow + 1, 1, oldLastRow, 5);
    for (var row = 1; lastRow + row <= oldLastRow; row++) {
      ledger.getRange(lastRow + row, 1, 1, 5).copyTo(ledger.getRange(row+1, 1, 1, 5));
    }
    remainingLedgerRange.clearContent();
  }
}
