#!/usr/bin/env node
import 'dotenv/config';
import { google } from 'googleapis';
import pLimit from 'p-limit';

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
const TEMPLATE_PRESENTATION_ID = '1YrKAl9DlHncNcP-ZxQMvuH8RO4Sbwx-jL0zfeUd9pHM';

const DATE_STR     = new Date().toISOString().slice(0, 10);
const CONCURRENCY  = 2;
const MAX_RETRIES  = 5;

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

// 🔥 Recoger gráficos de Dashboard en adelante
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

async function createTempPresentation(name, parentId) {
  const file = await withRetry('drive.copy presentation', () =>
    driveApi.files.copy({
      fileId: TEMPLATE_PRESENTATION_ID,
      requestBody: { name, parents: [parentId] },
      fields: 'id',
      supportsAllDrives: true,
    })
  );
  const presId = file.data.id;

  const pres = await withRetry('slides.get', () =>
    slidesApi.presentations.get({ presentationId: presId })
  );

  const pgW = pres.data.pageSize?.width?.magnitude || 960;
  const pgH = pres.data.pageSize?.height?.magnitude || 540;

  // 🗑️ borrar primera slide (portada de plantilla)
  const firstSlideId = pres.data.slides?.[0]?.objectId;
  if (firstSlideId) {
    await withRetry('slides.batchUpdate:deleteFirstSlide', () =>
      slidesApi.presentations.batchUpdate({
        presentationId: presId,
        requestBody: {
          requests: [{ deleteObject: { objectId: firstSlideId } }]
        }
      })
    );
  }

  return { presId, pgW, pgH };
}

// Crear nueva slide en blanco
async function createNewSlide(presId) {
  const slideId = `slide_${Date.now()}_${Math.random().toString(36).substr(2,5)}`;
  await withRetry('slides.batchUpdate:createSlide', () =>
    slidesApi.presentations.batchUpdate({
      presentationId: presId,
      requestBody: {
        requests: [
          {
            createSlide: {
              objectId: slideId,
              slideLayoutReference: { predefinedLayout: 'BLANK' }
            }
          }
        ]
      }
    })
  );
  return slideId;
}

// Insertar gráfico como objeto interactivo y escalar proporcionalmente
async function insertChartAndFit({ presId, slideId, chartId, pgW, pgH }) {
  const chartElemId = `chart_${chartId}_${Date.now()}`;

  // 1. Crear gráfico interactivo SIN size/transform → evita UNIT_UNSPECIFIED
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
              elementProperties: { pageObjectId: slideId }
            }
          }
        ]
      }
    })
  );

  // 2. Leer tamaño real del gráfico insertado
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

  // 3. Escalado proporcional y centrado
  const scale = Math.min(pgW / elemW, pgH / elemH);
  const offsetX = (pgW - elemW * scale) / 2;
  const offsetY = (pgH - elemH * scale) / 2;

  // 4. Aplicar transform
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
                scaleX: scale,
                scaleY: scale,
                shearX: 0,
                shearY: 0,
                translateX: offsetX,
                translateY: offsetY,
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

    const { presId, pgW, pgH } = await createTempPresentation(`${tienda}__${DATE_STR}`, dateFolderId);

    await Promise.all(charts.map((c, i) => limit(async () => {
      const idx = i + 1;
      const title = (c.title || `Grafico_${idx}`).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);

      try {
        const slideId = await createNewSlide(presId);
        await insertChartAndFit({ presId, slideId, chartId: c.chartId, pgW, pgH });
        console.log(`📊 OK ${tienda} gráfico ${idx} (${title})`);
        total++;
        await sleep(200);
      } catch (e) {
        console.log(`❌ Falló ${tienda} chart#${idx} (${title}): ${e.message || e}`);
      }
    })));
  }

  console.log(`✅ Presentaciones creadas con gráficos: ${total}`);
}

main().catch(err => { console.error('💥 Error:', err); process.exit(1); });
