const { google } = require('googleapis');
const sheets = google.sheets('v4');
const credentials = require('./desafioapi-413423-32a7aa0853ad.json'); // Replace with the correct path to your JSON credentials file

/**
 * Main function to process student data in the Google Sheets.
 */
async function main() {
  const spreadsheetId = '1haw_Nuq-iSiPDceuD3IX_0brGNhAPv7zqqTjms2b584';

  // Set up authentication using credentials
  const auth = new google.auth.GoogleAuth({
    credentials,
    scopes: ['https://www.googleapis.com/auth/spreadsheets'],
  });

  // Create Google Sheets client
  const sheetsClient = await sheets.spreadsheets.get({
    auth,
    spreadsheetId,
  });

  const sheet = sheetsClient.data.sheets[0];

  // Get data from the spreadsheet
  const response = await sheets.spreadsheets.values.get({
    auth,
    spreadsheetId,
    range: sheet.properties.title, // Assuming you want to work with the first sheet
  });

  const values = response.data.values;

  // Process student data
  for (let i = 3; i < values.length; i++) {
    const grade1 = parseFloat(values[i][3]);
    const grade2 = parseFloat(values[i][4]);
    const grade3 = parseFloat(values[i][5]);
    const absences = parseInt(values[i][2]);

    const average = (grade1 + grade2 + grade3) / 3;

    let situation;
    if (absences / 60 > 0.25) {
      situation = 'Failed due to Absences';
    } else if (average < 5) {
      situation = 'Failed due to Grade';
    } else if (average < 7) {
      situation = 'Final Exam';
      const finalExamGrade = Math.ceil((10 - average) * 2);
      values[i][7] = finalExamGrade.toString(); // Column for "Final Exam Grade"
    } else {
      situation = 'Approved';
      values[i][7] = '0'; // Set "Final Exam Grade" to 0 for students who are not in "Final Exam"
    }

    values[i][6] = situation; // Column for "Situation"
    console.log(`Student ${values[i][1]}: ${situation}`);
  }

  // Write back to the spreadsheet without changing the header row
  await sheets.spreadsheets.values.update({
    auth,
    spreadsheetId,
    range: sheet.properties.title,
    valueInputOption: 'RAW',
    resource: { values },
  });

  console.log('Process completed successfully!');
}

// Execute the main function
main();
