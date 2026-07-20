function recordAllFunds() {

  const funds = [
    {
      sheet: "Sheet Name",
      url: "url to what to track"
    },
    {
      sheet: "Sheet 2",
      url: "url 2"
    }
  ];

  const SPREADSHEET_ID = "YOUR_SPREADSHEET_ID";
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);

  funds.forEach(fund => {
    updateFund(ss, fund);
  });

}


function updateFund(ss, fund) {

  let sheet = ss.getSheetByName(fund.sheet);

  if (!sheet) {
    sheet = ss.insertSheet(fund.sheet);
    sheet.appendRow([
      "Date",
      "NAV",
      "Daily Change",
      "% Change"
    ]);
  }

  const response = UrlFetchApp.fetch(fund.url, {
    headers: {
      "User-Agent": "Mozilla/5.0"
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() != 200) {
    Logger.log("Failed to fetch " + fund.sheet);
    return;
  }

  const html = response.getContentText();

  const match = html.match(/\$([0-9]+\.[0-9]{2})/);

  if (!match) {
    Logger.log("Couldn't find NAV for " + fund.sheet);
    return;
  }

  const nav = parseFloat(match[1]);

  const today = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy-MM-dd"
  );

  const lastRow = sheet.getLastRow();

  // Don't record twice in one day
  if (lastRow > 1) {

    const lastDate = Utilities.formatDate(
      new Date(sheet.getRange(lastRow,1).getValue()),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd"
    );

    if (lastDate === today) {
      Logger.log(fund.sheet + " already updated today.");
      return;
    }
  }

  let change = "";
  let pct = "";

  if (lastRow > 1) {

    const previousNAV = Number(sheet.getRange(lastRow,2).getValue());

    change = nav - previousNAV;
    pct = change / previousNAV;
  }

  sheet.appendRow([
    new Date(),
    nav,
    change,
    pct
  ]);

  sheet.getRange(sheet.getLastRow(),4).setNumberFormat("0.00%");

  Logger.log("Updated " + fund.sheet + " : " + nav);

}
