eval(UrlFetchApp.fetch('https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.18.1/moment.min.js').getContentText());

function sortBalancesByDate(balances) {
  balances.sort(compareBalances); 
}

function compareBalances(balance1, balance2) {
  const date1 = balance1.date;
  const date2 = balance2.date;
  if(date1.isBefore(date2)) {
    return -1;
  } else if(date1.isAfter(date2)) {
    return 1;
  } else {
    return balance1.row - balance2.row;
  }
}

function convertToMoment(date) {
  return moment(date).startOf('day'); 
}

function getEndOfMonth(date) {
  return moment(date).endOf('month').startOf('day'); 
}

function getMonthRange(year, month) {
  const beginning = moment({year: year, month: month-1}).startOf('day');
  return {
    start: beginning,
    end: getEndOfMonth(beginning)
  };
}
