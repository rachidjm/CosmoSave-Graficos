#!/usr/bin/env node
import 'dotenv/config';
import { google } from 'googleapis';
import pLimit from 'p-limit';
import { Readable } from 'stream';

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
if (!SPREADSHEET_ID) { console.error('❌ Falta SPREADSHEET_ID'); process.exit(1); }

// --- OAuth de USUARIO ---
const OAUTH_CLIENT_ID = process.env.GOOGLE_OAUTH_CLIENT_ID;
const OAUTH_CLIENT_SECRET = process.env.GOOGLE_OAUTH_CLIENT_SECRET;
const OAUTH_REFRESH_TOKEN = process.env.GOOGLE_OAUTH_REFRESH_TOKEN;
if (!OAUTH_CLIENT_ID || !OAUTH_CLIENT_SECRET || !OAUTH_REFRESH_TOKEN) {
  console.error('❌ Faltan variables OAuth: GOOGLE_OAUTH_CLIENT_ID / GOOGLE_OAUTH_CLIENT_SECRET / GOOGLE_OAUTH_REFRESH_TOKEN');
  process.exit(1);
}
const oAuth2Client = new google.auth.OAuth2(
  OAUTH_CLIENT_ID,
  OAUTH_CLIENT_SECRET,
  'http://localhost:3000/oauth2callback'
);
oAuth2Client.setCredentials({ refresh_token: OAUTH_REFRESH_TOKEN });

const auth = oAuth2Client;
const sheetsApi = google.sheets({ version: 'v4', auth });
const driveApi  = google.drive({ version: 'v3', auth });
const slidesApi = google.slides({ version: 'v1', auth });

const TIENDAS = {
  ARENAL:          { sheetName: 'Dashboard',           folderId: '16PALsypZSdXiiXIgA_xMex710usAZAAZ' },
  DRUNI:           { sheetName: 'Dashboard D',         folderId: '1GrDRvmo9lR0RaBIw6y69OdFGV4Ao3KGi' },
  PRIETO:          { sheetName: 'Dashboard P',         folderId: '1mLoqIfnAb8QCqFlXrRciWBjGr43zMCCI' },
  AROMAS:          { sheetName: 'Dashboard A',         folderId: '1wXoQ4X3Ot2FYDGhDQ2v_c9BImd5SQNVv' },
  MARVIMUNDO:      { sheetName: 'Dashboard M',         folderId: '1jzHtaKBm2yMbLrnDCF6v9N8DPO8fmYZh' },
  JULIA:           { sheetName: 'Dashboard J',         folderId: '15Bn8zn26RW_2YTqMwVI4gPGx_1V4WQi4' },
  PACOPERFUMERIAS: { sheetName: 'Dashboard PF',        folderId: '1AtdZilQVDTJvFe1T09z102XQNZK8O49J' },
  PERSONALES:      { sheetName: 'GRAFICOS PERSONALES', folderId: '1cwLOPdclOxy47Bkp7dwvhzHLIIjB4svO' },
};

// 📂 Carpeta PTC en tu Drive personal
const TEMP_FOLDER_ID = '18vTs2um4CCqnI1OKWfBdM5_bnqLSeSJO';
// 🧩 ID de la plantilla en PTC
const TEMPLATE_PRESENTATION_ID = '1YrKAl9DlHncNcP-ZxQMvuH8RO4Sbwx-jL0zfeUd9pHM';

const FILE_PREFIX  = 'Grafico';
const DATE_STR     = new Date().toISOString().slice(0, 10);
const CONCURRENCY  = 2;
const MAX_RETRIES  = 5;

// 📏 Margen eliminado (pantalla completa)
const MARGIN_PT = 0;

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
async function withRetry(tag, fn) {
  let wait = 700;
  for (let i = 1; i <= MAX_RETRIES; i++) {
    try { return await fn(); }
    catch (e) {
      if (i === MAX_RETRIES) throw new Error(`${tag}: ${e?.message || e}`);
      console.log(`↻ Retry ${i}/${MAX_RETRIES} ${tag} en ${wait}ms: ${e?.message || e}`);
      await sleep(wait + Math.floor(Math.random() * 300));
      wait = Math.min(wait * 2, 8000);
    }
  }
}

async function getSheetsAndCharts() {
  const fields = 'sheets(properties(sheetId,title),charts(chartId,spec(title)))';
  const res = await withRetry('sheets.get', () =>
    sheetsApi.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields })
  );

  const byTitle = new Map();
  (res.data.sheets || []).forEach(sh => {
    const title = sh.properties?.title;
    if (!title) return;
    if (!title.startsWith('Dashboard') && title !== 'GRAFICOS PERSONALES') return;

    const sheetId = sh.properties?.sheetId;
    const charts = (sh.charts || []).map(c => ({
      chartId: c.chartId,
      title: c.spec?.title || ''
    }));

    if (charts.length) byTitle.set(title, { sheetId, title, charts });
  });

  return byTitle;
}

async function ensureDatedSubfolder(parentId, dateStr) {
  const q = `name='${dateStr}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`;
  const found = await withRetry('drive.list datedFolder', () =>
    driveApi.files.list({ q, fields: 'files(id,name)', spaces: 'drive', pageSize: 1, supportsAllDrives: true })
  );
  if (found.data.files?.length) return found.data.files[0].id;

  const folder = await withRetry('drive.create datedFolder', () =>
    driveApi.files.create({
      requestBody: { name: dateStr, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] },
      fields: 'id,name',
      supportsAllDrives: true,
    })
  );
  return folder.data.id;
}

async function createTempPresentation(name) {
  const file = await withRetry('drive.copy presentation', () =>
    driveApi.files.copy({
      fileId: TEMPLATE_PRESENTATION_ID,
      requestBody: { name, parents: [TEMP_FOLDER_ID] },
      fields: 'id',
      supportsAllDrives: true,
    })
  );
  const presId = file.data.id;

  const pres = await withRetry('slides.get', () =>
    slidesApi.presentations.get({ presentationId: presId })
  );

  const slideId = pres.data.slides?.[0]?.objectId;
  const pgW = pres.data.pageSize?.width?.magnitude || 960;
  const pgH = pres.data.pageSize?.height?.magnitude || 540;
  if (!slideId) throw new Error('No se pudo obtener slideId inicial');
  return { presId, slideId, pgW, pgH };
}

async function insertChartAndFit({ presId, slideId, chartId, pgW, pgH }) {
  const chartElemId = `chart_${chartId}_${Date.now()}`;

  await withRetry('slides.batchUpdate:createChart', () =>
    slidesApi.presentations.batchUpdate({
      presentationId: presId,
      requestBody: {
        requests: [
          {
            createSheetsChart: {
              objectId: chartElemId,
              spreadsheetId: SPREADSHEET_ID,
              chartId: chartId,
              linkingMode: 'LINKED',
              elementProperties: {
                pageObjectId: slideId,
                size: {
                  height: { magnitude: pgH, unit: 'PT' },
                  width:  { magnitude: pgW, unit: 'PT' }
                },
                transform: {
                  scaleX: 1,
                  scaleY: 1,
                  shearX: 0,
                  shearY: 0,
                  translateX: 0,
                  translateY: 0,
                  unit: 'PT'
                }
              }
            }
          }
        ]
      }
    })
  );

  const pres = await withRetry('slides.get after insert', () =>
    slidesApi.presentations.get({
      presentationId: presId,
      fields: 'slides(pageElements(objectId,size))'
    })
  );
  const elem = pres.data.slides
    .flatMap(s => s.pageElements || [])
    .find(e => e.objectId === chartElemId);

  const elemW = elem?.size?.width?.magnitude || 100;
  const elemH = elem?.size?.height?.magnitude || 100;

  const margin = MARGIN_PT;
  const targetW = pgW - 2 * margin;
  const targetH = pgH - 2 * margin;
  const scaleX = targetW / elemW;
  const scaleY = targetH / elemH;

  const translateX = (pgW - elemW * scaleX) / 2;
  const translateY = (pgH - elemH * scaleY) / 2;

  await withRetry('slides.batchUpdate:fit', () =>
    slidesApi.presentations.batchUpdate({
      presentationId: presId,
      requestBody: {
        requests: [
          {
            updatePageElementTransform: {
              objectId: chartElemId,
              applyMode: 'ABSOLUTE',
              transform: {
                scaleX,
                scaleY,
                shearX: 0,
                shearY: 0,
                translateX,
                translateY,
                unit: 'PT'
              }
            }
          }
        ]
      }
    })
  );

  return chartElemId;
}

async function exportPresentationPDF(presId) {
  const res = await withRetry('drive.export(pdf)', () =>
    driveApi.files.export(
      { fileId: presId, mimeType: 'application/pdf' },
      { responseType: 'stream' }
    )
  );

  const chunks = [];
  return await new Promise((resolve, reject) => {
    res.data.on('data', chunk => chunks.push(chunk));
    res.data.on('end', () => resolve(Buffer.concat(chunks)));
    res.data.on('error', reject);
  });
}

async function deletePageElement(presId, objectId) {
  await withRetry('slides.batchUpdate:deleteElement', () =>
    slidesApi.presentations.batchUpdate({
      presentationId: presId,
      requestBody: { requests: [{ deleteObject: { objectId } }] }
    })
  );
}

function bufferToStream(buffer) {
  return new Readable({
    read() {
      this.push(buffer);
      this.push(null);
    }
  });
}

async function uploadPDF({ parentId, name, pdfBuffer }) {
  await withRetry(`drive.upload ${name}`, () =>
    driveApi.files.create({
      requestBody: { name, parents: [parentId], mimeType: 'application/pdf' },
      media: { mimeType: 'application/pdf', body: bufferToStream(pdfBuffer) },
      fields: 'id',
      supportsAllDrives: true,
    })
  );
}

async function main() {
  const { token } = await auth.getAccessToken();
  if (!token) { console.error('❌ No se pudo obtener access token OAuth'); process.exit(1); }

  const byTitle = await getSheetsAndCharts();
  const limit = pLimit(CONCURRENCY);
  let total = 0;

  for (const [tienda, { sheetName, folderId }] of Object.entries(TIENDAS)) {
    const sh = byTitle.get(sheetName);
    if (!sh) { console.log(`⚠️ Hoja no encontrada: "${sheetName}" (${tienda})`); continue; }
    const charts = sh.charts || [];
    if (!charts.length) { console.log(`ℹ️ ${tienda} / ${sheetName}: sin gráficos incrustados`); continue; }

    let dateFolderId;
    try { dateFolderId = await ensureDatedSubfolder(folderId, DATE_STR); }
    catch (e) { console.log(`❌ Carpeta destino de ${tienda} inválida: ${e.message || e}`); continue; }

    console.log(`🗂️ ${tienda} / ${sheetName}: ${charts.length} gráficos → ${DATE_STR}`);

    const { presId, slideId, pgW, pgH } = await createTempPresentation(`TMP_${tienda}__${DATE_STR}`);

    await Promise.all(charts.map((c, i) => limit(async () => {
      const idx = i + 1;
      const title = (c.title || `${FILE_PREFIX}_${idx}`).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
      const fileName = `${tienda}__${title}__${DATE_STR}.pdf`;

      try {
        const objId = await insertChartAndFit({ presId, slideId, chartId: c.chartId, pgW, pgH });
        const pdf = await exportPresentationPDF(presId);
        await uploadPDF({ parentId: dateFolderId, name: fileName, pdfBuffer: pdf });
        await deletePageElement(presId, objId);

        console.log(`📄 OK ${tienda} → ${fileName}`);
        total++;
        await sleep(600);   // ⏳ antes eran 200 ms → ahora 600 ms
      } catch (e) {
        console.log(`❌ Falló ${tienda} chart#${idx} (${title}): ${e.message || e}`);
      }
    })));

    try {
      await withRetry('drive.delete pres', () =>
        driveApi.files.delete({ fileId: presId, supportsAllDrives: true })
      );
    } catch (e) {
      console.log(`⚠️ No se pudo borrar presentación temporal de ${tienda}: ${e.message || e}`);
    }
  }

  console.log(`✅ Export completado. PDFs subidos: ${total}`);
}

main().catch(err => { console.error('💥 Error:', err); process.exit(1); });
