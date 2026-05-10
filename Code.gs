/**
 * Google Apps Script — Auto Insurance Lead Form webhook
 * Deploy as Web app (Execute as: Me, Who has access: Anyone)
 */
function doPost(e) {
  var data = JSON.parse(e.postData.contents);
  data.timestamp = new Date().toISOString();
  var cols = [
    'timestamp',
    'fullName',
    'email',
    'phone',
    'zip',
    'dob',
    'sex',
    'maritalStatus',
    'occupation',
    'education',
    'vin',
    'vehicle',
    'primaryUse',
    'householdResidents',
    'priorCarrier',
    'policyExpiration',
    'effectiveDate',
    'coverageBI',
    'coveragePD',
    'coverageUM',
    'coverageComp',
    'coverageColl',
    'coverageRental',
    'coverageTowing',
    'authorized'
  ];
  var row = cols.map(function (k) {
    var v = data[k];
    return v === undefined || v === null ? '' : v;
  });
  SpreadsheetApp.getActiveSpreadsheet().getActiveSheet().appendRow(row);
  return ContentService.createTextOutput(JSON.stringify({ status: 'ok' })).setMimeType(
    ContentService.MimeType.JSON
  );
}
