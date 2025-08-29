#!/usr/bin/env node
import { google } from 'googleapis';

const SCOPES = ['https://www.googleapis.com/auth/drive'];
const auth = new google.auth.GoogleAuth({ scopes: SCOPES });
const drive = google.drive({ version: 'v3', auth });

async function purgeTrash() {
  console.log('🚮 Vaciando la papelera de la Service Account...');
  await drive.files.emptyTrash();
  console.log('✅ Papelera vaciada por completo');
}

purgeTrash().catch(err => {
  console.error('💥 Error al vaciar la papelera:', err.message);
  process.exit(1);
});
