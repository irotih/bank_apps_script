/**
 * Get the current interest APR
 */
function getCurrentInterestRate() {
    return getLedger().getRange(Constants.InterestRateLocation.row, Constants.InterestRateLocation.column).getValue();
}

/**
 * Calculate the daily accruals between 2 arbitrary dates assuming no balance change between the dates
 */
function calculateDailyAccrual(balance, startDate, endDate, dailyRate) {
    const numDays = getDaysDiff(endDate, startDate);
    const accrual = numDays * balance * dailyRate;
    Logger.log('Calculating Daily Accrual. Start date: ' +
              startDate.toString() + ', End date: ' +
              endDate.toString() + ', No. Days: ' +
              numDays + ', Balance: ' + balance + ', Rate: ' +
              dailyRate + ', Accrual: ' + accrual);
    return accrual;
}

/**
 * Calculate the interest amount for a single month
 */
function calculateInterest(endOfLastMonth, endingBalance, endOfThisMonth, lineItems) {
    var interest = 0.0;
    var dailyRate = getCurrentInterestRate() / 365;
    sortBalancesByDate(lineItems);
    var lastDate = endOfLastMonth;
    var startingBalance = endingBalance;
    for(var i = 0; i<lineItems.length; i++) {
        if(i < lineItems.length - 1 && isSame(lineItems[i].date, lineItems[i+1].date)) {
            //Skip the row if next row has the same date. This guarantees we get the latest ending balance
            continue;
        }
        interest += calculateDailyAccrual(startingBalance, lastDate, lineItems[i].date, dailyRate);
        lastDate = lineItems[i].date;
        startingBalance = lineItems[i].balance;
    }
    //Calculate interest for the remainer of the month
    if (isBefore(lastDate, endOfThisMonth)) {
        interest += calculateDailyAccrual(startingBalance, lastDate, endOfThisMonth, dailyRate);
    } 
    return Math.round(interest * 100) / 100;
}

/**
 * Return the last interest as
 * {
 *    date: date of interest payment,
 *    row: row number of interest row
 * }
 */
function getLastInterest() {
    const range = getLedger().getDataRange().getValues();
    for(var row = range.length - 1; row >= 0; row--) {
        if(range[row][1] === 'Interest') {
            return {
                date: convertDateToMidnight(range[row][0]),
                row: row+1
            };
        }
    }
    return null;
}

/**
 * Get the last date in each month that has no interest payments
 */
function getMissingInterestMonths(firstDateOfAccrual) {
    var missingMonths = [];
    var currentMonth = getStartOfMonth(new Date());
    var lastDateOfMonth = getEndOfMonth(firstDateOfAccrual);
    while(isBefore(lastDateOfMonth, currentMonth)) {
        missingMonths.push(lastDateOfMonth);
        lastDateOfMonth = getEndOfNextMonth(lastDateOfMonth);
    }
    return missingMonths;
}

/**
 * Creates interest payments lines for every month up until
 * the last month for any months that don't have interest
 * posted. Balances are updated if interest is posted
 * before existing transactions.
 */
function generateInterestPayments() {
    const ledger = getLedger();
    if(isEmpty(ledger)) {
        return;
    }
    const lastInterest = getLastInterest();
    var firstDateOfAccrual;
    var lastBalance;
    if(lastInterest) {
        firstDateOfAccrual = getStartOfNextMonth(lastInterest.date);
        lastBalance = getLastBalanceOfMonth(lastInterest.date);
    } else {
        firstDateOfAccrual = getStartingDate(ledger);
        lastBalance = getStartingBalance(ledger);
    }
    const missingMonths = getMissingInterestMonths(firstDateOfAccrual);
    var lineItems = [];
    var interest = 0;
    var insertionRow = 0;
    for(var i=0; i<missingMonths.length; i++) {
        lineItems = getLinesForMonth(missingMonths[i].getFullYear(), missingMonths[i].getMonth()+1);
        interest = calculateInterest(firstDateOfAccrual, lastBalance, missingMonths[i], lineItems);
        insertionRow = getInsertionRow(missingMonths[i]);
        if(insertionRow <= ledger.getLastRow()) {
            shiftLinesDown(insertionRow);
        }
        ledger.getRange(insertionRow, 1, 1, 3).setValues([[missingMonths[i], 'Interest', interest]]);
        recalculateBalances(insertionRow - 1);
        firstDateOfAccrual = missingMonths[i];
        lastBalance = getLedger().getRange(insertionRow, 4).getValue();
    }
}