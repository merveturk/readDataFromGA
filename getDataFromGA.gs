
  function readData(profileId) {
    var lastr=mainshett.getLastRow();
    var startdate= mainshett.getRange(lastr,1).getValue();
    var enddate= mainshett.getRange(lastr,2).getValue();
    var startDate = Utilities.formatDate(startdate, Session.getScriptTimeZone(),'yyyy-MM-dd');
    var endDate = Utilities.formatDate(enddate, Session.getScriptTimeZone(),'yyyy-MM-dd');

    var tableId = 'ga:****';
    var metric = 'ga:entrances';
    var options = {
      'dimensions': 'ga:pagePath',
      'sort': 'ga:source',
      'filters': 'ga:pagePath=~^/fr',
      'filters': 'ga:sourceMedium==google / organic',
      'max-results': 25
    };
    var report = Analytics.Data.Ga.get(tableId, startDate, endDate, metric,
        options);

    if (report.rows) {

      var sheet=SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
      //var spreadsheet = SpreadsheetApp.create('Google Analytics Report');
      //var sheet = spreadsheet.getActiveSheet();

      // Append the headers.
      var headers = report.columnHeaders.map(function(columnHeader) {
        //return columnHeader.name;
      });
      //sheet.appendRow(headers);

      // Append the results.
      sheet.getRange(lastr, 32, report.rows.length, headers.length)
          .setValues(report.rows);

      //Logger.log('Report spreadsheet created: %s',
        //  spreadsheet.getUrl());
    } else {
      Logger.log('No rows returned.');
    }
  }
