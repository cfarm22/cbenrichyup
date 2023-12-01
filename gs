function enrichEmails() {
  // Replace 'YOUR_CLEARBIT_API_KEY' with your actual Clearbit API key
  var clearbitApiKey = 'sk_2880d21c9becfaf373b6fd9b16ee8201';
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var emails = sheet.getRange('A:A').getValues(); // Assuming emails are in column A

  for (var i = 1; i <= emails.length; i++) {
    var email = emails[i - 1][0];

    if (validateEmail(email)) {
      // Call Clearbit API
      var clearbitResponse = getClearbitData(email, clearbitApiKey);

      // Update the sheet with Clearbit response
      updateSheetWithClearbitData(sheet, i, clearbitResponse);
    } else {
      // Handle invalid email
      sheet.getRange('B' + i).setValue('Invalid Email');
    }
  }
}

function validateEmail(email) {
  // Simple email validation using a regular expression
  var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return emailRegex.test(email);
}

function getClearbitData(email, apiKey) {
  var url = 'https://person.clearbit.com/v2/combined/find?email=' + encodeURIComponent(email);
  var options = {
    'headers': {
      'Authorization': 'Bearer ' + apiKey
    }
  };

  try {
    var response = UrlFetchApp.fetch(url, options);

    // Handle invalid email response
    if (response.getResponseCode() === 422) {
      var errorResponse = JSON.parse(response.getContentText());
      return 'Clearbit Error: ' + errorResponse.error.message;
    }

    var data = JSON.parse(response.getContentText());
    return data; // Return the raw Clearbit data
  } catch (error) {
    // Handle other errors
    return 'Error: ' + error.toString();
  }
}

function updateSheetWithClearbitData(sheet, row, clearbitResponse) {
  // Assuming clearbitResponse is an object with various fields
  // Adjust the following lines based on the actual structure of the Clearbit response

  if (clearbitResponse && clearbitResponse.person) {
    sheet.getRange('B' + row).setValue(clearbitResponse.person.name && clearbitResponse.person.name.fullName || '');
    sheet.getRange('C' + row).setValue(clearbitResponse.person.email && clearbitResponse.person.email.address || '');
    sheet.getRange('D' + row).setValue(clearbitResponse.person.location && clearbitResponse.person.location || '');
    // Add more lines for other fields you want to extract
  } else {
    sheet.getRange('B' + row).setValue('No Clearbit Data');
  }
}

