// Google Sheets API Service using service account
const { google } = require('googleapis');

// Service account credentials
const SERVICE_ACCOUNT_EMAIL = 'google-sheets-editor@ai-chat-widget-adminpanel.iam.gserviceaccount.com';
const SERVICE_ACCOUNT_KEY = require('./service-account-key.json');

// Spreadsheet ID and sheet name
const SPREADSHEET_ID = '1pjdQl0Uf798HRT9C1QlfAHPKBdaN9lzdrHgu6FTDxso';
const SHEET_NAME = 'Sheet1';

// Initialize Google Sheets API client
let sheets;
let isInitialized = false;

async function initializeSheets() {
    try {
        const auth = new google.auth.GoogleAuth({
            keyFile: './service-account-key.json',
            scopes: ['https://www.googleapis.com/auth/spreadsheets']
        });
        const client = await auth.getClient();
        
        console.log('Successfully authenticated with Google Sheets API');
        sheets = google.sheets({ version: 'v4', auth: client });
        isInitialized = true;
    } catch (err) {
        console.error('Error initializing Google Sheets API:', err);
        throw new Error('Failed to initialize Google Sheets API client');
    }
}

// Initialize on first use
async function ensureInitialized() {
    if (!isInitialized) {
        await initializeSheets();
    }
}

/**
 * Appends form data to Google Sheet
 * @param {Object} formData - Form data to append
 * @returns {Promise} Promise resolving when data is appended
 */
async function appendToSheet(formData) {
    await ensureInitialized();
    try {
        console.log('Appending form data to sheet:', formData);
        
        const response = await sheets.spreadsheets.values.append({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!A:K`,
            valueInputOption: 'USER_ENTERED',
            insertDataOption: 'INSERT_ROWS',
            resource: {
                values: [[
                    formData.idNumber,
                    formData.isMissing ? 'נעדר' : '',
                    formData.name,
                    formData.street,
                    formData.houseNumber,
                    formData.apartmentNumber,
                    formData.floor,
                    formData.residents,
                    formData.familyMissing ? 'כן' : 'לא',
                    formData.clothing || '',
                    formData.additionalFamily || '',
                    formData.mobilePhone,
                    formData.emergencyPhone,
                    formData.homePhone,
                    new Date().toLocaleString('he-IL')
                ]]
            }
        });
        
        console.log('Successfully appended data to sheet');
        return response;
    } catch (error) {
        console.error('Error appending to sheet:', error);
        throw new Error(`Failed to append data to sheet: ${error.message}`);
    }
}

/**
 * Gets all records from sheet and filters by ID number
 * @param {string} idNumber - ID number to search for
 * @returns {Promise} Promise resolving with matching record or null
 */
async function getRecordById(idNumber) {
    await ensureInitialized();
    try {
        console.log(`Looking up record for ID: ${idNumber}`);
        
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!A:K`
        });

        const rows = response.data.values;
        if (!rows || rows.length === 0) {
            console.log('No data found in sheet');
            return null;
        }

        // Find matching record (assuming ID is in first column)
        const record = rows.find(row => row[0] && row[0].toString().trim() === idNumber);
        
        if (record) {
            console.log('Found matching record:', record);
            return {
                idNumber: record[0],
                name: record[1],
                street: record[2],
                houseNumber: record[3],
                apartmentNumber: record[4],
                floor: record[5],
                mobilePhone: record[6],
                emergencyPhone: record[7],
                homePhone: record[8]
            };
        }
        
        console.log('No matching record found');
        return null;
    } catch (error) {
        console.error('Error getting record:', error);
        throw new Error(`Failed to lookup record: ${error.message}`);
    }
}

async function testReadSheet() {
    await ensureInitialized();
    try {
        const response = await sheets.spreadsheets.values.get({
            spreadsheetId: SPREADSHEET_ID,
            range: `${SHEET_NAME}!A:K`
        });
        
        console.log('Successfully read sheet data:');
        console.log('First 5 rows:', response.data.values.slice(0, 5));
        return response.data.values;
    } catch (error) {
        console.error('Error reading sheet:', error);
        throw error;
    }
}

module.exports = { appendToSheet, getRecordById, testReadSheet };
