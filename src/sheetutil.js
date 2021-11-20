/**
 * @returns The Ledger sheet
 */
function getLedger() {
    return SpreadsheetApp.getActive().getSheetByName(Constants.SheetName.LEDGER);
}

/**
 * @returns The History sheet 
 */
function getHistory() {
    return SpreadsheetApp.getActive().getSheetByName(Constants.SheetName.HISTORY);
}

/**
 * Checks if the specified row is completely filled (required fields)
 * @param sheet 
 * @param row 
 * @returns 
 */
function isRowComplete(sheet, row) {
    const rowValues = sheet.getRange(row, 1, 1, Constants.NUM_REQ_DATA_COLUMNS).getValues();
    for (var i = 0; i < Constants.NUM_REQ_DATA_COLUMNS; i++) {
        if (rowValues[0][i] === '') {
            return false;
        }
    }
    return true;
}

/**
 * Checks if the specified row is completely empty
 * @param sheet 
 * @param row 
 * @returns 
 */
function isRowEmpty(sheet, row) {
    const rowValues = sheet.getRange(row, 1, 1, Constants.NUM_DATA_COLUMNS).getValues();
    for (var i = 0; i < Constants.NUM_DATA_COLUMNS; i++) {
        if (rowValues[0][i] !== '') {
            return false;
        }
    }
    return true;
}

function getFirstDataRow(sheet) {
    switch(sheet.getName()) {
        case 'History':
            return Constants.HistoryRow.FIRST_DATA_ROW;
        case 'Ledger':
            return Constants.LedgerRow.FIRST_DATA_ROW;
        default:
            throw new Error('Unknown sheet named '+sheet.getName());
    }
}

/**
 * Checks if the sheet is empty (no data rows)
 * @param sheet 
 * @returns 
 */
function isEmpty(sheet) {
    return sheet.getLastRow() === 1 || isRowEmpty(sheet, getFirstDataRow(sheet));
}

/**
 * Get the last row id containing data
 * @param sheet 
 * @throws If the last data row is not complete or if the sheet has no data rows
 */
function getLastDataRow(sheet) {
    const lastSheetRow = sheet.getLastRow();
    if (lastSheetRow === 1) {
        throw new Error('Sheet is empty');
    }
    for (var row = lastSheetRow; row >= getFirstDataRow(sheet); row--) {
        if (isRowComplete(sheet, row)) {
            return row;
        } else if (!isEmpty(sheet, row)) {
            throw new Error('Last row is not complete');
        }
    }
    throw new Error('Sheet is empty');
}

/**
 * Get the final balance for the provided sheet
 * @param sheet 
 * @returns 
 */
function getSheetBalance(sheet) {
    if (isEmpty(sheet)) {
        return 0;
    }
    const lastRow = getLastDataRow(sheet);
    return sheet.getRange(lastRow, Constants.Column.BALANCE).getValue();
}

/**
 * Get the final current balance.
 * @returns 
 */
function getCurrentBalance() {
    return getSheetBalance(getLedger());
}

/**
 * Get the last balance in History tab
 * @returns 
 */
function getHistoricBalance() {
    return getSheetBalance(getHistory());
}

/**
 * Get the starting date on the provided sheet
 * @returns 
 */
function getStartingDate(sheet) {
    if (isEmpty(sheet)) {
        return null;
    }
    return convertDateToMidnight(sheet.getRange(getFirstDataRow(sheet), Constants.Column.DATE).getValue());
}

/**
 * Get the ending date on the provided sheet
 * @returns 
 */
function getEndingDate(sheet) {
    if (isEmpty(sheet)) {
        return null;
    }
    return convertDateToMidnight(sheet.getRange(getLastDataRow(sheet), Constants.Column.DATE).getValue());
}

/**
 * Get the starting balance on the provided sheet
 * @returns 
 */
function getStartingBalance(sheet) {
    if (isEmpty(sheet)) {
        return null;
    }
    return sheet.getRange(getFirstDataRow(sheet), Constants.Column.BALANCE).getValue();
}

/**
 * Get the ending balance on the provided sheet
 * @param sheet 
 * @returns 
 */
function getEndingBalance(sheet) {
    if (isEmpty(sheet)) {
        return null;
    }
    return sheet.getRange(getLastDataRow(sheet), Constants.Column.BALANCE).getValue();
}

function shiftRowsUp(sheet, firstRowToShift, lastRowToShift, destinationRow) {
    const numRowsToShift = lastRowToShift - firstRowToShift + 1;
    const rangeToShift = sheet.getRange(firstRowToShift, 1, numRowsToShift, Constants.NUM_DATA_COLUMNS);
    rangeToShift.copyTo(sheet.getRange(destinationRow, 1, numRowsToShift, Constants.NUM_DATA_COLUMNS));

    const firstLeftover = destinationRow + numRowsToShift;
    const leftoverRows = lastRowToShift - firstLeftover + 1;
    sheet.getRange(firstLeftover, 1, leftoverRows, Constants.NUM_DATA_COLUMNS).clearContent();
}
