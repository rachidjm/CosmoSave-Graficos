#!/usr/bin/env node
/**
 * Exporta todos los gráficos incrustados de pestañas indicadas a PDF
 * y los sube a Drive en subcarpetas YYYY-MM-DD.
 *
 * npm i googleapis pdfkit node-fetch@3 p-limit image-size dotenv
 *
 * Secrets en GitHub Actions:
 *   - SPREADSHEET_ID             (ID del Google Sheet)
 *   - SA_JSON                    (JSON del Service Account entero)
 *
 * IMPORTANTE: Comparte el Spreadsheet y las carpetas destino de Drive
 * con el email del Service Account (Editor).
 */

import 'dotenv/config';
import fs from 'node:fs';
import fetch from 'node-fetch';
import PDFDocument from 'pdfkit';
import { google } from 'googleapis';
import pLimit from 'p-limit';
import { imageSize } from 'image-size';

/* =======================
 *  CONFIG
 * ======================= */
const SPREADSHEET_ID = process.env.SPREADSHEET_ID;
if (!SPREADSHEET_ID) {
  console.error('❌ Falta SPREADSHEET_ID (secret)'); process.exit(1);
}

// Mapa tienda -> { sheetName, folderId }
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
const DATE_STR     = new Date().toISOString().slice(0, 10);
const CONCURRENCY  = 3;
const MAX_RETRIES  = 5;

/* =======================
 *  AUTH (Service Account)
 * ======================= */
const SCOPES = [
  'https://www.googleapis.com/auth/spreadsheets.readonly',
  'https://www.googleapis.com/auth/drive',
  'https://www.googleapis.com/auth/drive.file',
];

let CREDENTIALS;
try {
  CREDENTIALS = JSON.parse(process.env.SA_JSON || '');
} catch (e) {
  console.error('❌ SA_JSON no es JSON válido'); process.exit(1);
}

const auth = new google.auth.GoogleAuth({ credentials: CREDENTIALS, scopes: SCOPES });
const sheetsApi = google.sheets({ version: 'v4', auth });
const driveApi  = google.drive({ version: 'v3', auth });

/* =======================
 *  HELPERS
 * ======================= */
const sleep = (ms) => new Promise(r => setTimeout(r, ms));

async function withRetry(tag, fn) {
  let wait = 600;
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

async function downloadChartPNG({ sheetId, chartId, accessToken }) {
  const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=png&gid=${sheetId}&oid=${chartId}`;
  const res = await withRetry(`fetch chart gid=${sheetId} oid=${chartId}`, () =>
    fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } })
  );
  if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.length < 1000) throw new Error('PNG demasiado pequeño');
  return buf;
}

function pngToPDFBuffer(pngBuffer) {
  const { width, height } = imageSize(pngBuffer) || {};
  const isLandscape = (width && height) ? width >= height : true;
  const pdf = new PDFDocument({ size: 'A4', layout: isLandscape ? 'landscape' : 'portrait', autoFirstPage: false });
  const chunks = [];
  pdf.on('data', d => chunks.push(d));
  const done = new Promise(resolve => pdf.on('end', () => resolve(Buffer.concat(chunks))));
  pdf.addPage();
  const page = pdf.page;
  const maxW = page.width  - page.margins.left - page.margins.right;
  const maxH = page.height - page.margins.top  - page.margins.bottom;
  let drawW = maxW, drawH = maxH;
  if (width && height) {
    const scale = Math.min(maxW / width, maxH / height);
    drawW = Math.floor(width * scale);
    drawH = Math.floor(height * scale);
  }
  const x = page.margins.left + (maxW - drawW) / 2;
  const y = page.margins.top  + (maxH - drawH) / 2;
  pdf.image(pngBuffer, x, y, { width: drawW, height: drawH });
  pdf.end();
  return done;
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

async function uploadPDF({ parentId, name, pdfBuffer }) {
  await withRetry(`drive.upload ${name}`, () =>
    driveApi.files.create({
      requestBody: { name, parents: [parentId], mimeType: 'application/pdf' },
      media: { mimeType: 'application/pdf', body: Buffer.from(pdfBuffer) },
      fields: 'id',
    })
  );
}

/* =======================
 *  MAIN
 * ======================= */
async function main() {
  const accessToken = await auth.getAccessToken();
  if (!accessToken) { console.error('❌ Sin accessToken'); process.exit(1); }

  const byTitle = await getSheetsAndCharts();
  const limit = pLimit(CONCURRENCY);
  let total = 0;

  for (const [tienda, { sheetName, folderId }] of Object.entries(TIENDAS)) {
    const sh = byTitle.get(sheetName);
    if (!sh) { console.log(`⚠️ Hoja no encontrada: ${sheetName} (${tienda})`); continue; }
    if (!sh.charts?.length) { console.log(`ℹ️ ${tienda}/${sheetName}: sin gráficos incrustados`); continue; }

    let dateFolderId;
    try { dateFolderId = await ensureDatedSubfolder(folderId, DATE_STR); }
    catch (e) { console.log(`❌ Carpeta destino ${tienda} inválida: ${e.message || e}`); continue; }

    console.log(`🗂️ ${tienda} / ${sheetName}: ${sh.charts.length} gráficos → ${DATE_STR}`);

    const tasks = sh.charts.map((c, i) => limit(async () => {
      const idx = i + 1;
      const title = (c.title || `${FILE_PREFIX}_${idx}`).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
      const fileName = `${tienda}__${title}__${DATE_STR}.pdf`;
      try {
        const png = await downloadChartPNG({ sheetId: sh.sheetId, chartId: c.chartId, accessToken });
        const pdf = await pngToPDFBuffer(png);
        await uploadPDF({ parentId: dateFolderId, name: fileName, pdfBuffer: pdf });
        console.log(`📄 OK ${tienda} → ${fileName}`);
        total++;
      } catch (e) {
        console.log(`❌ Falló ${tienda} chart#${idx} (${title}): ${e.message || e}`);
      }
    }));

    await Promise.all(tasks);
  }

  console.log(`✅ Export completado. PDFs subidos: ${total}`);
}

main().catch(err => { console.error('💥 Error:', err); process.exit(1); });
