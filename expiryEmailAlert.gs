// This script will take something's expiry date from the seventh column of a Google sheet and check to see if it is expiring today, in 2 weeks, or in 1 month.
// It will then send an automated email to the provided email address alerting the impending expiration date.

// NB: If your expiry date is in a different column than column 6 (they start counting from 0), update the column numbers in lines 61 & 74.

// NB: To limit the time spent searching, limit the active number of searched rows in line 59 to the rows of active data - code also assumes only one row of headers in line 55.

// NB: Replace the email@gmail.com email with the address you wish to notify on lines 146, 178, and 207. Email message can be edited on lines 89-119; email subject is set in lines 144, 170, & 199.

function expiryEmailAlert() {

  // Set today's date information

  var today = new Date();

  var todayMonth = today.getMonth() + 1;

  var todayDay = today.getDate();

  var todayYear = today.getFullYear();

  
  // Set the date 2 weeks from now

  var twoWeeksFromToday = new Date();

  twoWeeksFromToday.setDate(twoWeeksFromToday.getDate() + 14);

  var twoWeeksMonth = twoWeeksFromToday.getMonth() + 1;

  var twoWeeksDay = twoWeeksFromToday.getDate();

  var twoWeeksYear = twoWeeksFromToday.getFullYear();


  // Set the date 1 month from now

  var newToday = new Date()

  var oneMonthFromToday = new Date(newToday.setMonth(newToday.getMonth()+1));

  var oneMonthMonth = oneMonthFromToday.getMonth() + 1;

  var oneMonthDay = oneMonthFromToday.getDate();

  var oneMonthYear = oneMonthFromToday.getFullYear();

 
  // Find date from spreadsheet

  var sheet = SpreadsheetApp.getActiveSheet();

  // First row of data to process

  var startRow = 2;

  // Number of rows to process

  var numRows = 51;
 
  var dataRange = sheet.getRange(startRow, 1, numRows, 7);

  var data = dataRange.getValues();

 
  // Loop through all of the rows

  for (var i = 0; i < data.length; ++i) {

    var row = data[i];

    var expireDateFormat = Utilities.formatDate(

      new Date(row[6]),

      'ET',

      'MM/dd/yyyy'

    );

 
    // Email message information

    var subject = '';

    var message =

      ' A Halal certificate is expiring. ' +

      '\n' +

      ' Code: ' +

      row[0] +

      '\n' +

      ' Description: ' +

      row[2] +

      '\n' +

      ' Supplier: ' +

      row[3] +

      '\n' +

      ' Issued: ' +

      row[5] +

      '\n' +

      ' Expiration: ' +

      expireDateFormat;

    // Expiry date information

    var expireDateMonth = new Date(row[6]).getMonth() + 1;

    var expireDateDay = new Date(row[6]).getDate();

    var expireDateYear = new Date(row[6]).getFullYear();

 
    // Check for things expiring today & send alert email

    if (

      expireDateMonth === todayMonth &&

      expireDateDay === todayDay &&

      expireDateYear === todayYear

    ) {

      var subject =

        'A Halal Certificate expired today: ' + row[0] + ' - ' + expireDateFormat;

      MailApp.sendEmail('email@gmail.com', subject, message);

      Logger.log('todayyyy!');

    }


    // Check for things expiring 2 weeks from now & send alert email

 
    Logger.log('2 weeks month, expire month' + twoWeeksMonth + expireDateMonth);

    if (

      expireDateMonth === twoWeeksMonth &&

      expireDateDay === twoWeeksDay &&

      expireDateYear === twoWeeksYear

    ) {

      var subject =

        'A Halal Certificate is expiring in 2 weeks: ' +

        row[0] +

        ' - ' +

        expireDateFormat;

      MailApp.sendEmail('email@gmail.com', subject, message);

      Logger.log('2 weeks from now');

    }

 
    // Check for things expiring 1 month from now & send alert email

    if (

      expireDateMonth === oneMonthMonth &&

      expireDateDay === oneMonthDay &&

      expireDateYear === oneMonthYear

    ) {

      var subject =

        'A Halal Certificate is expiring in 1 month: ' +

        row[0] +

        ' - ' +

        expireDateFormat;

      MailApp.sendEmail('email@gmail.com', subject, message);

      Logger.log('1 month from now');

    }

  }

}
