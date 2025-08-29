#!/usr/bin/env node
import { google } from 'googleapis';

const SCOPES = ['https://www.googleapis.com/auth/drive'];
const auth = new google.auth.GoogleAuth({ scopes: SCOPES });
const drive = google.drive({ version: 'v3', auth });

async function cleanupAll() {
  console.log('🧹 Buscando TODOS los archivos de la Service Account...');

  let pageToken = null;
  let count = 0;

  do {
    const res = await drive.files.list({
      q: "trashed=false", // solo archivos activos
      fields: "nextPageToken, files(id, name, mimeType)",
      pageSize: 1000,
      pageToken,
    });

    const files = res.data.files || [];
    if (files.length === 0 && !pageToken) {
      console.log('✅ No hay archivos en la cuenta, ya está limpia.');
      return;
    }

    for (const f of files) {
      try {
        await drive.files.delete({ fileId: f.id });
        console.log(`🗑️ Borrado: ${f.name} (${f.mimeType})`);
        count++;
      } catch (e) {
        console.log(`⚠️ No se pudo borrar ${f.name}: ${e.message}`);
      }
    }

    pageToken = res.data.nextPageToken;
  } while (pageToken);

  console.log(`✅ Limpieza completada. Archivos eliminados: ${count}`);
}

cleanupAll().catch(err => {
  console.error('💥 Error en cleanup:', err.message);
  process.exit(1);
});
