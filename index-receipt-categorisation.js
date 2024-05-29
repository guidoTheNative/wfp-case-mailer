// Author: "JOHN CHALERA <john.chalera@wfp.org>"
const fs = require('fs');
const { google } = require('googleapis');
const nodemailer = require('nodemailer');

// Load client secrets from a local file.
const credentials = require('./service-account.json'); // Path to your service account JSON file

// Email mappings based on priority and district
const emailMappings = {
    high_priority: 'high_priority@example.com',
    medium_priority: 'medium_priority@example.com',
    low_priority: 'low_priority@example.com',
    districts: {
        'District A': 'districtA@example.com',
        'District B': 'districtB@example.com',
        'District C': 'districtC@example.com'
    },
    both: {
        'high_priority_districtA@example.com': { priority: 'High', district: 'District A' },
        'medium_priority_districtB@example.com': { priority: 'Medium', district: 'District B' },
        'low_priority_districtC@example.com': { priority: 'Low', district: 'District C' }
    }
};

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
 * Watch the Google Sheet for changes.
 * @param {google.auth.JWT} auth The authenticated Google JWT client.
 */
function watchSpreadsheet(auth) {
    const sheets = google.sheets({ version: 'v4', auth });
    const spreadsheetId = '1HkI7_IXf-ZD80pnzOHowecJ4ncUXlB418Q_R5fmu5fI';
    const range = 'Sheet1!A:Z'; // Adjust the range as per your sheet structure

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
                        headers.forEach((header, index) => {
                            const value = entry[index] !== undefined ? entry[index] : 'N/A';
                            formattedText += `<li><strong>${header}:</strong> ${value}</li>`;
                        });
                        formattedText += '</ul>';
                        
                        const priority = entry[1]; // Assuming priority is in the second column
                        const district = entry[2]; // Assuming district is in the third column
                        const recipients = getRecipients(priority, district);
                        
                        sendEmail(subject, formattedText, recipients);
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

    // Check for changes every 10 seconds
    setInterval(checkForChanges, 10000);
}

/**
 * Get the list of recipients based on priority and district.
 * @param {string} priority The priority level.
 * @param {string} district The district.
 * @returns {string} Comma-separated list of recipient emails.
 */
function getRecipients(priority, district) {
    let recipients = [];

    // Add priority-based email
    if (priority) {
        const priorityEmail = emailMappings[`${priority.toLowerCase()}_priority`];
        if (priorityEmail) {
            recipients.push(priorityEmail);
        }
    }

    // Add district-based email
    if (district) {
        const districtEmail = emailMappings.districts[district];
        if (districtEmail) {
            recipients.push(districtEmail);
        }
    }

    // Add combined priority and district-based email
    for (const email in emailMappings.both) {
        const mapping = emailMappings.both[email];
        if (mapping.priority === priority && mapping.district === district) {
            recipients.push(email);
        }
    }

    return recipients.join(',');
}

/**
 * Send an email notification.
 * @param {string} subject The subject of the email.
 * @param {string} html The body of the email in HTML format.
 * @param {string} recipients Comma-separated list of recipient emails.
 */
function sendEmail(subject, html, recipients) {
    const transporter = nodemailer.createTransport({
        service: 'gmail',
        auth: {
            user: 'mailing.immalawi@gmail.com',
            pass: 'jxqcbsqugfjysdyz'
        },
    });

    let mailOptions = {
        from: 'mailing.immalawi@gmail.com',
        to: recipients,
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
