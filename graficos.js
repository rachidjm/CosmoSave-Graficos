#!/usr/bin/env node
import 'dotenv/config';
import { google } from 'googleapis';
import pLimit from 'p-limit';

/**
 * Exporta TODOS los gráficos incrustados de las pestañas indicadas:
 * 1) Crea UNA presentación temporal en Slides
 * 2) Inserta el gráfico desde Sheets (createSheetsChart)
 * 3) Ajusta a toda la página
 * 4) Exporta a PDF con Drive.files.export
 * 5) Sube el PDF a la carpeta destino/YYYY-MM-DD en Drive
 * 6) Reutiliza la MISMA slide borrando el elemento anterior (rápido y sin cuotas tontas)
 *
 * Requiere en el workflow:
 *  - GOOGLE_APPLICATION_CREDENTIALS → sa.json (ya lo tienes)
 *  - secret SPREADSHEET_ID
 * Imprescindible: habilitar **Slides API** en el proyecto del Service Account.
 */

const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
if (!SPREADSHEET_ID) { console.error('❌ Falta SPREADSHEET_ID'); process.exit(1); }

// Mapa tienda -> { sheetName, folderId }  (tus IDs reales)
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

const FILE_PREFIX  = 'Grafico';
const DATE_STR     = new Date().toISOString().slice(0, 10); // YYYY-MM-DD
const CONCURRENCY  = 2;   // inserción+export PDF es más pesado; 2 es seguro
const MAX_RETRIES  = 5;

const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets.readonly',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/drive.file',
  'https://www.googleapis.com/auth/presentations', // Slides API
];

const auth = new google.auth.GoogleAuth({ scopes: SCOPES });
const sheetsApi = google.sheets({ version: 'v4', auth });
const driveApi  = google.drive({ version: 'v3', auth });
const slidesApi = google.slides({ version: 'v1', auth });

const sleep = (ms) => new Promise(r => setTimeout(r, ms));
async function withRetry(tag, fn) {
  let wait = 700;
  for (let i = 1; i <= MAX_RETRIES; i++) {
    try { return await fn(); }
    catch (e) {
      if (i === MAX_RETRIES) throw new Error(`${tag}: ${e.message || e}`);
      console.log(`↻ Retry ${i}/${MAX_RETRIES} ${tag} en ${wait}ms: ${e.message || e}`);
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
    const sheetId = sh.properties?.sheetId;
    const charts = (sh.charts || []).map(c => ({ chartId: c.chartId, title: c.spec?.title || '' }));
    if (title) byTitle.set(title, { sheetId, title, charts });
  });
  return byTitle;
}

async function ensureDatedSubfolder(parentId, dateStr) {
  const q = `name='${dateStr}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`;
  const found = await withRetry('drive.list datedFolder', () =>
    driveApi.files.list({ q, fields: 'files(id,name)', spaces: 'drive', pageSize: 1 })
  );
  if (found.data.files?.length) return found.data.files[0].id;

  const folder = await withRetry('drive.create datedFolder', () =>
    driveApi.files.create({
      requestBody: { name: dateStr, mimeType: 'application/vnd.google-apps.folder', parents: [parentId] },
      fields: 'id,name',
    })
  );
  return folder.data.id;
}

/** Crea presentación temporal con una slide en blanco y devuelve info de tamaño. */
async function createTempPresentation(name) {
  const pres = await withRetry('slides.create', () =>
    slidesApi.presentations.create({ requestBody: { title: name } })
  );
  const presId = pres.data.presentationId;
  const slideId = pres.data.slides?.[0]?.objectId;
  const pgW = pres.data.pageSize?.width?.magnitude || 960;
  const pgH = pres.data.pageSize?.height?.magnitude || 540;
  if (!slideId) throw new Error('No se pudo obtener slideId inicial');
  return { presId, slideId, pgW, pgH };
}

/** Inserta un Sheets Chart en la slide (como elemento enlazado) y lo ajusta a toda la página. */
async function insertChartAndFit({ presId, slideId, chartId, pgW, pgH }) {
  const chartElemId = `chart_${chartId}_${Date.now()}`;
  const requests = [
    {
      createSheetsChart: {
        objectId: chartElemId,
        spreadsheetId: SPREADSHEET_ID,
        chartId,
        linkingMode: 'LINKED'
      }
    },
    {
      // Mover el chart a la slide destino
      insertSlidesObject: {
        objectId: chartElemId,
        slideObjectId: slideId
      }
    }
  ];

  await withRetry('slides.batchUpdate:createChart', () =>
    slidesApi.presentations.batchUpdate({
      presentationId: presId,
      requestBody: { requests }
    })
  );

  // Tamaño y posición para ocupar toda la página (con pequeños márgenes)
  const margin = 10;
  await withRetry('slides.batchUpdate:fit', () =>
    slidesApi.presentations.batchUpdate({
      presentationId: presId,
      requestBody: {
        requests: [
          {
            updateSize: {
              objectId: chartElemId,
              size: {
                height: { magnitude: pgH - 2 * margin, unit: 'PT' },
                width:  { magnitude: pgW - 2 * margin, unit: 'PT' }
              },
              fields: 'height,width'
            }
          },
          {
            updatePageElementTransform: {
              objectId: chartElemId,
              applyMode: 'ABSOLUTE',
              transform: {
                scaleX: 1, scaleY: 1, shearX: 0, shearY: 0,
                translateX: margin, translateY: margin, unit: 'PT'
              }
            }
          }
        ]
      }
    })
  );

  return chartElemId;
}

/** Exporta la presentación a PDF (una página) y devuelve Buffer. */
async function exportPresentationPDF(presId) {
  const res = await withRetry('drive.export(pdf)', () =>
    driveApi.files.export({ fileId: presId, mimeType: 'application/pdf' }, { responseType: 'arraybuffer' })
  );
  return Buffer.from(res.data);
}

/** Borra el elemento de la slide para reutilizarla con el siguiente chart. */
async function deletePageElement(presId, objectId) {
  await withRetry('slides.batchUpdate:deleteElement', () =>
    slidesApi.presentations.batchUpdate({
      presentationId: presId,
      requestBody: { requests: [{ deleteObject: { objectId } }] }
    })
  );
}

/** Envía el PDF a la carpeta destino. */
async function uploadPDF({ parentId, name, pdfBuffer }) {
  await withRetry(`drive.upload ${name}`, () =>
    driveApi.files.create({
      requestBody: { name, parents: [parentId], mimeType: 'application/pdf' },
      media: { mimeType: 'application/pdf', body: Buffer.from(pdfBuffer) },
      fields: 'id',
    })
  );
}

async function main() {
  // Comprobación mínima de acceso (token)
  const client = await auth.getClient();
  const token = await client.getAccessToken();
  if (!token) { console.error('❌ No se pudo obtener token'); process.exit(1); }

  // Mapa título -> charts
  const byTitle = await getSheetsAndCharts();

  const limit = pLimit(CONCURRENCY);
  let total = 0;

  for (const [tienda, { sheetName, folderId }] of Object.entries(TIENDAS)) {
    const sh = byTitle.get(sheetName);
    if (!sh) { console.log(`⚠️ Hoja no encontrada: "${sheetName}" (${tienda})`); continue; }
    const charts = sh.charts || [];
    if (!charts.length) { console.log(`ℹ️ ${tienda} / ${sheetName}: sin gráficos incrustados`); continue; }

    // Subcarpeta por fecha
    let dateFolderId;
    try { dateFolderId = await ensureDatedSubfolder(folderId, DATE_STR); }
    catch (e) { console.log(`❌ Carpeta destino de ${tienda} inválida: ${e.message || e}`); continue; }

    console.log(`🗂️ ${tienda} / ${sheetName}: ${charts.length} gráficos → ${DATE_STR}`);

    // 1 presentación temporal por tienda
    const { presId, slideId, pgW, pgH } = await createTempPresentation(`TMP_${tienda}__${DATE_STR}`);

    // Procesamos gráficos (limit de concurrencia bajo)
    await Promise.all(charts.map((c, i) => limit(async () => {
      const idx = i + 1;
      const title = (c.title || `${FILE_PREFIX}_${idx}`).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
      const fileName = `${tienda}__${title}__${DATE_STR}.pdf`;

      try {
        // Inserta chart enlazado y ajusta
        const objId = await insertChartAndFit({ presId, slideId, chartId: c.chartId, pgW, pgH });

        // Exporta la presentación a PDF (1 página)
        const pdf = await exportPresentationPDF(presId);

        // Sube PDF a Drive
        await uploadPDF({ parentId: dateFolderId, name: fileName, pdfBuffer: pdf });

        // Elimina el elemento para reutilizar la slide
        await deletePageElement(presId, objId);

        console.log(`📄 OK ${tienda} → ${fileName}`);
        total++;
        await sleep(200); // respiro
      } catch (e) {
        console.log(`❌ Falló ${tienda} chart#${idx} (${title}): ${e.message || e}`);
      }
    })));

    // Envía la presentación a la papelera
    try {
      await withRetry('drive.trash pres', () =>
        driveApi.files.update({ fileId: presId, requestBody: { trashed: true } })
      );
    } catch (e) {
      console.log(`⚠️ No se pudo eliminar presentación temporal de ${tienda}: ${e.message || e}`);
    }
  }

  console.log(`✅ Export completado. PDFs subidos: ${total}`);
}

main().catch(err => { console.error('💥 Error:', err); process.exit(1); });
