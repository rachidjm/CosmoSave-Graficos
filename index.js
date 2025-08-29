require('dotenv').config();
const { google } = require('googleapis');
const { crearGraficosParaTodas } = require('./graficos');

async function main() {
  const creds = JSON.parse(process.env.SHEETS_PRIVATE_KEY || '{}');

  const auth = new google.auth.GoogleAuth({
    credentials: {
      client_email: creds.client_email,
      private_key: (creds.private_key || '').replace(/\\n/g, '\n')
    },
    scopes: ['https://www.googleapis.com/auth/spreadsheets']
  });

  await crearGraficosParaTodas({
    auth,
    spreadsheetId: process.env.SPREADSHEET_ID
  });
}

main().catch(e => { console.error('❌', e.message); process.exit(1); });
