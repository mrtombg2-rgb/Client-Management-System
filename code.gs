const SHEETS = {
  CLIENTS: "Clients",
  PRODUCTS: "Products",
  TRANSACTIONS: "Transactions",
  PAYMENTS: "Payments",
  DASHBOARD: "Dashboard",
  SETTINGS: "Settings"
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("CMS Pro")
    .addItem("Refresh Dashboard", "refreshDashboard")
    .addItem("Health Check", "healthCheck")
    .addToUi();
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const name = sheet.getName();
  const row = e.range.getRow();
  if (row < 2) return;

  if (name === SHEETS.TRANSACTIONS) {
    handleTransactionEdit_(e);
  }

  if (name === SHEETS.PAYMENTS) {
    handlePaymentEdit_(e);
  }
}

function handleTransactionEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();

  if (!sh.getRange(row, 1).getValue()) {
    sh.getRange(row, 1).setValue("TRX" + Utilities.formatString("%06d", row - 1));
  }

  sh.getRange(row, 13).setValue(new Date());

  if (e.range.getColumn() === 4) {
    fillProductPrice_(sh, row);
  }

  calculateTotals_(sh, row);
}

function handlePaymentEdit_(e) {
  const sh = e.range.getSheet();
  const row = e.range.getRow();

  if (!sh.getRange(row, 1).getValue()) {
    sh.getRange(row, 1).setValue("PAY" + Utilities.formatString("%06d", row - 1));
  }

  sh.getRange(row, 9).setValue(new Date());
}

function fillProductPrice_(sheet, row) {
  const product = sheet.getRange(row, 4).getValue();
  const ps = SpreadsheetApp.getActive().getSheetByName(SHEETS.PRODUCTS);
  if (!ps) return;

  const lastRow = ps.getLastRow();
  if (lastRow < 2) return;

  const data = ps.getRange(2, 2, lastRow - 1, 3).getValues();
  for (const d of data) {
    if (d[0] === product) {
      sheet.getRange(row, 7).setValue(d[1]);
      sheet.getRange(row, 8).setValue(d[2]);
      break;
    }
  }
}

function calculateTotals_(sheet, row) {
  const qty = Number(sheet.getRange(row, 6).getValue()) || 0;
  const unitUsd = Number(sheet.getRange(row, 7).getValue()) || 0;
  const unitBdt = Number(sheet.getRange(row, 8).getValue()) || 0;

  sheet.getRange(row, 9).setValue(qty * unitUsd);
  sheet.getRange(row, 10).setValue(qty * unitBdt);
}

function refreshDashboard() {
  const ss = SpreadsheetApp.getActive();
  const dash = ss.getSheetByName(SHEETS.DASHBOARD);
  if (!dash) {
    SpreadsheetApp.getUi().alert("Dashboard sheet not found.");
    return;
  }

  const trx = ss.getSheetByName(SHEETS.TRANSACTIONS);
  const pay = ss.getSheetByName(SHEETS.PAYMENTS);

  let salesUsd = 0;
  let salesBdt = 0;
  let paidUsd = 0;

  if (trx && trx.getLastRow() >= 2) {
    const tData = trx.getRange(2, 1, trx.getLastRow() - 1, Math.max(trx.getLastColumn(), 10)).getValues();
    tData.forEach(r => {
      salesUsd += Number(r[8]) || 0;
      salesBdt += Number(r[9]) || 0;
    });
  }

  if (pay && pay.getLastRow() >= 2) {
    const pData = pay.getRange(2, 1, pay.getLastRow() - 1, Math.max(pay.getLastColumn(), 5)).getValues();
    pData.forEach(r => {
      paidUsd += Number(r[3]) || 0;
    });
  }

  dash.getRange("B2").setValue(salesUsd);
  dash.getRange("B3").setValue(salesBdt);
  dash.getRange("B4").setValue(paidUsd);
  dash.getRange("B5").setValue(salesUsd - paidUsd);
}

function healthCheck() {
  const ss = SpreadsheetApp.getActive();
  const missing = [];

  Object.values(SHEETS).forEach(name => {
    if (!ss.getSheetByName(name)) missing.push(name);
  });

  if (missing.length) {
    SpreadsheetApp.getUi().alert("Missing sheets: " + missing.join(", "));
  } else {
    SpreadsheetApp.getUi().alert("CMS structure looks OK.");
  }
}
