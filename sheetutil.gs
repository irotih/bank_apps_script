/**
 * @returns The Ledger sheet
 */
function getLedger() {
    return SpreadsheetApp.getActive().getSheetByName(SheetName.LEDGER);
}

/**
 * @returns The History sheet 
 */
function getHistory() {
    return SpreadsheetApp.getActive().getSheetByName(SheetName.HISTORY);
}

/**
 * Checks if the specified row is completely filled (required fields)
 * @param sheet 
 * @param row 
 * @returns 
 */
function isRowComplete(sheet, row) {
    const rowValues = sheet.getRange(row, 1, 1, NUM_REQ_DATA_COLUMNS).getValues();
    for (let value of rowValues) {
        if (value === '') {
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
    const rowValues = sheet.getRange(row, 1, 1, NUM_DATA_COLUMNS).getValues();
    for (let value of rowValues) {
        if (value !== '') {
            return false;
        }
    }
    return true;
}

/**
 * Checks if the sheet is empty (no data rows)
 * @param sheet 
 * @returns 
 */
function isEmpty(sheet) {
    return sheet.getLastRow() === 1 || isRowEmpty(sheet, Row.FIRST_DATA_ROW);
}

/**
 * Get the last row id containing data
 * @param sheet 
 * @throws If the last data row is not complete or if the sheet has no data rows
 */
function getLastDataRow(sheet) {
    const lastSheetRow = sheet.getLastDataRow();
    if (lastSheetRow === 1) {
        throw new Error('Sheet is empty');
    }
    for (let row = lastSheetRow; row >= Row.FIRST_DATA_ROW; row--) {
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
    return sheet.getRange(lastRow, Column.BALANCE).getValue();
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
    return convertToMoment(sheet.getRange(Row.FIRST_DATA_ROW, Column.DATE).getValue());
}

/**
 * Get the ending date on the provided sheet
 * @returns 
 */
function getEndingDate(sheet) {
    if (isEmpty(sheet)) {
        return null;
    }
    return convertToMoment(sheet.getRange(getLastDataRow(sheet), Column.DATE).getValue());
}

/**
 * Get the starting balance on the provided sheet
 * @returns 
 */
function getStartingBalance(sheet) {
    if (isEmpty(sheet)) {
        return null;
    }
    return sheet.getRange(Row.FIRST_DATA_ROW, Column.BALANCE).getValue();
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
    return sheet.getRange(getLastDataRow(sheet), Column.BALANCE).getValue();
}
