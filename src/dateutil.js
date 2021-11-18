/**
 * @param {Date} date1 
 * @param {Date} date2 
 * @returns true if date1 is before date2 
 */
function isBefore(date1, date2) {
  return date1.getTime() < date2.getTime();
}

/**
 * @param {Date} date1 
 * @param {Date} date2 
 * @returns true if date1 is after date2 
 */
function isAfter(date1, date2) {
  return date1.getTime() > date2.getTime();
}

function compareBalances(balance1, balance2) {
  const date1 = balance1.date;
  const date2 = balance2.date;
  if(isBefore(date1, date2)) {
    return -1;
  } else if(isAfter(date1, date2)) {
    return 1;
  } else {
    return balance1.row - balance2.row;
  }
}

function sortBalancesByDate(balances) {
  balances.sort(compareBalances); 
}

function convertDateToMidnight(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function isSame(date1, date2) {
  return convertDateToMidnight(date1).getTime() === convertDateToMidnight(date2).getTime();
}

/**
 * @param {Date} date 
 * @returns the last date in the month of the provided date
 */
function getEndOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth()+1, 0);
}

/**
 * @param {Date} date 
 * @returns the last date in the month prior to the provided date
 */
function getEndOfPriorMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 0);
}

/**
 * @param {Date} date 
 * @returns the last date in the month after the provided date
 */
function getEndOfNextMonth(date) {
  return new Date(date.getFullYear(), date.getMonth()+2, 0);
}

/**
 * @param {Date} date 
 * @returns the first date in the month of the provided date
 */
function getStartOfMonth(date) {
  return new Date(date.getFullYear(), date.getMonth(), 1);
}

/**
 * @param {*} year 
 * @param {*} month 1-12
 * @returns {
 *  // First date of the month
 *  start: Date;
 *  // Last date of the month
 *  end: Date;
 * }
 */
function getMonthRange(year, month) {
  const beginning = new Date(year, month-1, 1);
  return {
    start: beginning,
    end: getEndOfMonth(beginning)
  };
}

function getDaysDiff(date1, date2) {
  const diffTime = Math.abs(date2 - date1);
  return Math.ceil(diffTime / (1000 * 60 * 60 * 24)); 
}