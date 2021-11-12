function getLedger() {
  return SpreadsheetApp.getActive().getSheetByName('Ledger');
}

function getHistory() {
  return SpreadsheetApp.getActive().getSheetByName('History');
}

function isEmpty() {
  return getLedger().getLastRow() === 1 || getLedger().getRange(2,1).isBlank();
}

function getCurrentBalance() {
  return getLedger().getRange(2,8).getValue();
}

function getHistoricBalance() {
  return getHistory().getRange(1,8).getValue();
}

function getStartingDate() {
  if(isEmpty()) {
    return null;
  }
  return convertToMoment(getLedger().getRange(2,1).getValue());
}

function getEndingDate() {
  if(isEmpty()) {
    return null;
  }
  const ledger = getLedger();
  return convertToMoment(ledger.getRange(ledger.getLastRow(), 1).getValue());
}

function getStartingBalance() {
  if(isEmpty()) {
    return null;
  }
  return getLedger().getRange(2,4).getValue();
}

function getInsertionRow(date) {
  const ledger = getLedger();
  const range = ledger.getDataRange().getValues();
  for(var row = range.length - 1; row >= 1; row--) {
    if(convertToMoment(range[row][0]).isBefore(date)) {
      return row+2;
    }
  }
  return ledger.getLastRow();
}

function getLastBalanceOfMonth(date) {
  if(isEmpty()) {
    return 0;
  }
  const lastDateOfMonth = getEndOfMonth(date);
  const range = getLedger().getDataRange().getValues();
  var lastSeenBalance = range[range.length-1][3];
  for(var i = range.length; i < range.length; i++) {
    if(convertToMoment(range[i][0]).isAfter(date)) {
       return lastSeenBalance;
    }
    lastSeenBalance = range[i][3];
  }
  return lastSeenBalance;
}

function getLinesForMonth(year, month) {
  const dateRange = getMonthRange(year, month);
  const lines = [];
  const range = getLedger().getDataRange().getValues();
  var date = null;
  for(var i = 1; i<range.length; i++) {
    date = convertToMoment(range[i][0]);
    if(!(date.isBefore(dateRange.start) || date.isAfter(dateRange.end))) {
      lines.push({
        date: date,
        row: i+1,
        type: range[i][1],
        delta: range[i][2],
        balance: range[i][3]
      });
    }
    if(date.isAfter(dateRange.end)) {
      break;
    }
  }
  return lines;
}

function shiftLinesDown(firstRow) {
  const ledger = getLedger();
  const numRows = ledger.getLastRow() - firstRow + 1;
  const srcRange = ledger.getRange(firstRow, 1, numRows, 5);
  const destRange = ledger.getRange(firstRow+1, 1, numRows, 5);
  srcRange.copyTo(destRange);
  ledger.getRange(firstRow, 1, 1, 5).clearContent();
}
