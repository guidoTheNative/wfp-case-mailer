// Author: "JOHN CHALERA <john.chalera@wfp.org>"
const fs = require('fs');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');
const xlsx = require('xlsx');

const emailSheet = require('./emailSheet');


// Load client secrets from a local file.
const credentials = require('./service-account.json'); // Path to your service account JSON file


/**
 * Create a JWT client with the given credentials, and then execute the given callback function.
 * @param {Object} credentials The authorization client credentials.
 * @param {function} callback The callback to call with the authorized client.
 */
function authorize(credentials, callback) {
    const { client_email, private_key } = credentials;
    const jwtClient = new google.auth.JWT(
        client_email,
        null,
        private_key,
        ['https://www.googleapis.com/auth/spreadsheets.readonly']
    );
    callback(jwtClient);
}

/**
 * Get the latest month sheet name from Google Sheets.
 * @param {google.auth.JWT} auth The authenticated Google JWT client.
 * @param {string} spreadsheetId The ID of the spreadsheet.
 * @param {function} callback The callback to call with the sheet name.
 */
function getLatestMonthSheetName(auth, spreadsheetId, callback) {
    const sheets = google.sheets({ version: 'v4', auth });
    sheets.spreadsheets.get({ spreadsheetId }, (err, res) => {
        if (err) {
            console.error('Error retrieving spreadsheet information:', err);
            return;
        }

        const sheetNames = res.data.sheets.map(sheet => sheet.properties.title);
        const months = [
            "January", "February", "March", "April", "May", "June",
            "July", "August", "September", "October", "November", "December"
        ];
        const now = new Date();
        let month = now.getMonth();
        let year = now.getFullYear();

        while (month >= 0) {
            const sheetName = `${months[month]} ${year}`;
            if (sheetNames.includes(sheetName)) {
                callback(sheetName);
                return;
            }
            month--;
        }

        console.log('No sheet found for the current or previous months.');
        callback(null);
    });
}

/**
 * Watch the Google Sheet for changes.
 * @param {google.auth.JWT} auth The authenticated Google JWT client.
 */
function watchSpreadsheet(auth) {
    const spreadsheetId = '1mQtPBqIDHdkDRumLVF4XVxoR6Uln1063Pk_SbDHDD2A';
    getLatestMonthSheetName(auth, spreadsheetId, (latestSheetName) => {
        if (!latestSheetName) {
            console.log('No valid sheet found. Exiting.');
            return;
        }

        const sheets = google.sheets({ version: 'v4', auth });
        const range = `${latestSheetName}!A:Z`;

        let previousData = [];

        const checkForChanges = () => {
            sheets.spreadsheets.values.get({ spreadsheetId, range }, (err, res) => {
                if (err) {
                    console.log('The API returned an error: ' + err);
                    return;
                }
                const rows = res.data.values;
                if (rows && rows.length > 0) {
                    const headers = rows[0]; // First row as headers
                    const data = rows.slice(1); // Remaining rows as data

                    console.log('Previous data length: ', previousData.length);
                    console.log('Current data length: ', data.length);

                    if (previousData.length > 0 && data.length > previousData.length) {
                        console.log('New entries detected.');
                        const newEntries = data.slice(previousData.length);
                        console.log('New entries:', newEntries);

                        newEntries.forEach(entry => {
                            const subject = `New Case Reported: ${entry[0]}`;
                            let formattedText = '<p>A new case has been reported:</p><ul>';
                            let programme = 'Undefined';
                            let priority = 'Medium/Low';
                            let district = 'Undefined';

                            headers.forEach((header, index) => {
                                const value = entry[index] !== undefined ? entry[index] : 'N/A';
                                if (header === 'Programme') programme = value;
                                if (header === 'Priority') priority = value;
                                if (header === 'District') district = value;
                                formattedText += `<li><strong>${header}:</strong> ${value}</li>`;
                            });
                            formattedText += '</ul>';

                            // Determine recipient emails
                            let recipientEmails = determineRecipientEmails(programme, priority, district);
                            if (recipientEmails) { sendEmail(subject, formattedText, recipientEmails); }
                        });
                    } else {
                        if (previousData.length === 0) {
                            console.log('No previous data to compare.');
                        } else if (data.length <= previousData.length) {
                            console.log('No new entries detected.');
                        }
                    }
                    previousData = data;
                } else {
                    console.log('No data found.');
                }
            });
        };

        // Check for changes every 15 minutes
        setInterval(checkForChanges, 900000);
    });
}

/**
 * Determine recipient emails based on programme, priority, and district.
 * @param {string} programme The programme of the case.
 * @param {string} priority The priority of the case.
 * @param {string} district The district of the case.
 * @return {string} The recipient emails.
 */
/**
 * Determine recipient emails based on programme, priority, and district.
 * @param {string} programme The programme of the case.
 * @param {string} priority The priority of the case.
 * @param {string} district The district of the case.
 * @return {string|null} The recipient emails or null if no emails should be sent.
 */

const normalize = (str) => str ? str.trim().toLowerCase() : undefined;

const determineRecipientEmails = (programme, priority, district) => {
    const actualProgramme = normalize(programme);
    const normalizedPriority = normalize(priority);
    const normalizedDistrict = normalize(district);

    const uniqueEmails = new Set();
    emailSheet
        .filter(row => {
            const rowProgrammes = row.Programme.split(',').map(normalize);
            const rowPriorities = row.Priority.split('/').map(normalize);
            const rowDistricts = row.District.split(',').map(normalize);

            return (!actualProgramme || rowProgrammes.includes(actualProgramme)) &&
                (!normalizedPriority || rowPriorities.includes(normalizedPriority)) &&
                (!normalizedDistrict || rowDistricts.includes(normalizedDistrict));
        })
        .forEach(row => {
            const emails = row.Emails.split(';').map(email => email.trim());
            emails.forEach(email => uniqueEmails.add(email));
        });

    return [...uniqueEmails].join(';');
};

/**
 * Send an email notification.
 * @param {string} subject The subject of the email.
 * @param {string} html The body of the email in HTML format.
 * @param {string} recipientEmails The recipient emails.
 */
function sendEmail(subject, html, recipientEmails) {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'mailing.immalawi@gmail.com',
            pass: 'jxqcbsqugfjysdyz'
        },
    });

    let mailOptions = {
        from: 'mailing.immalawi@gmail.com',
        to: recipientEmails,
        subject: subject,
        html: html
    };

    transporter.sendMail(mailOptions, (error, info) => {
        if (error) {
            return console.log(error);
        }
        console.log('Email sent: ' + info.response);
    });
}

// Authorize and start watching the spreadsheet
authorize(credentials, watchSpreadsheet);
