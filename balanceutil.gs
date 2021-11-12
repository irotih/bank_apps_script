/**
 * Recalculate all balances from the provided starting row
 * onwards. If no row is provided, start at row 2.
 * Balances are only updated if their values differ from
 * what they should be based on the prior balance and
 * current activity. Updated balances will briefly flash
 * yellow.
 */
function recalculateBalances(startingRow) {
  if(isEmpty()) {
    return;
  }
  const range = getLedger().getDataRange();
  const lastRow = range.getLastRow();
  var i = startingRow && startingRow > 1 ? startingRow : 2;
  Logger.log('Recalculating balances starting at row '+i);
  var currentBalance = 0;
  var balanceCell = null;
  var newBalance = range.getCell(i++, 4).getValue();
  var delta = 0;
  const updatedCells = [];
  for(; i <= lastRow; i++) {
    delta = (range.getCell(i,2).getValue() === 'Withdrawal' ? -1 : 1 ) * range.getCell(i,3).getValue();
    newBalance += delta;
    balanceCell = range.getCell(i,4);
    currentBalance = balanceCell.getValue();
    if(currentBalance !== newBalance) {
      Logger.log('Updating balance for row '+i+' cell 4 from '+currentBalance+' to '+newBalance);
      updatedCells.push({ cell: balanceCell, background: balanceCell.getBackground() || 'white' });
      balanceCell.setValue(newBalance);
      balanceCell.setBackground('yellow');
    }
  }
  
  for(var j = 0; j < updatedCells.length; j++) {
    Logger.log('background: '+updatedCells[j].background);
    updatedCells[j].cell.setBackground(updatedCells[j].background);
  }
}
