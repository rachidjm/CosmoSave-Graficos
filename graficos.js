#!/usr/bin/env node
/**
 * Exporta todos los gráficos incrustados (Embedded Charts) de las pestañas indicadas
 * a PDF y los sube a Google Drive, dentro de subcarpetas por fecha (YYYY-MM-DD).
 *
 * Requisitos:
 *   - Node 18+
 *   - npm i googleapis pdfkit node-fetch@3 p-limit image-size dotenv
 *   - .env con:
 *       SPREADSHEET_ID=...
 *       SHEETS_PRIVATE_KEY={"type":"service_account",...}   // JSON del SA en una sola línea
 *
 * Importante:
 *   - Comparte el Spreadsheet y las carpetas DESTINO de Drive con el email del Service Account (Editor).
 */

import 'dotenv/config';
import fs from 'node:fs';
import path from 'node:path';
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
  console.error('❌ Falta SPREADSHEET_ID en .env');
  process.exit(1);
}

// Mapea tienda -> { sheetName, folderId }  (usa los tuyos)
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
  CREDENTIALS = JSON.parse(process.env.SHEETS_PRIVATE_KEY || '');
} catch (e) {
  console.error('❌ SHEETS_PRIVATE_KEY no es un JSON válido en .env');
  process.exit(1);
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
      if (i === MAX_RETRIES) throw new Error(`${tag}: agotados ${MAX_RETRIES} intentos → ${e.message || e}`);
      console.log(`↻ Retry ${i}/${MAX_RETRIES} ${tag} en ${wait}ms: ${e.message || e}`);
      await sleep(wait + Math.floor(Math.random() * 300));
      wait = Math.min(wait * 2, 8000);
    }
  }
}

/** Obtiene sheetId y chartIds de cada pestaña (solo charts incrustados, NO "hojas de gráfico"). */
async function getSheetsAndCharts() {
  const fields = 'sheets(properties(sheetId,title),charts(chartId,spec(title)))';
  const res = await withRetry('sheets.get', () =>
    sheetsApi.spreadsheets.get({ spreadsheetId: SPREADSHEET_ID, fields })
  );

  const data = res.data.sheets || [];
  const byTitle = new Map();
  data.forEach(sh => {
    const title = sh.properties?.title;
    const sheetId = sh.properties?.sheetId;
    const charts  = (sh.charts || []).map(c => ({
      chartId: c.chartId,
      title:   c.spec?.title || '',
    }));
    if (title) byTitle.set(title, { sheetId, title, charts });
  });
  return byTitle;
}

/** Descarga un chart incrustado como PNG usando la URL de export de Sheets. */
async function downloadChartPNG({ sheetId, chartId, accessToken }) {
  // URL "oficial de facto" de export de charts de Sheets (oid = chartId)
  const url = `https://docs.google.com/spreadsheets/d/${SPREADSHEET_ID}/export?format=png&gid=${sheetId}&oid=${chartId}`;
  const res = await withRetry(`fetch chart gid=${sheetId} oid=${chartId}`, () =>
    fetch(url, { headers: { Authorization: `Bearer ${accessToken}` } })
  );
  if (!res.ok) throw new Error(`HTTP ${res.status} ${res.statusText}`);
  const buf = Buffer.from(await res.arrayBuffer());
  if (buf.length < 1000) throw new Error('PNG demasiado pequeño (¿permiso o chart inexistente?)');
  return buf;
}

/** Convierte PNG a PDF (1 página), ajustando orientación y escala. */
function pngToPDFBuffer(pngBuffer, outName = 'chart.pdf') {
  const { width, height } = imageSize(pngBuffer) || {};
  const isLandscape = (width && height) ? width >= height : true;

  // A4 puntos: [595.28, 841.89]
  const pageSize = 'A4';
  const pdf = new PDFDocument({ size: pageSize, layout: isLandscape ? 'landscape' : 'portrait', autoFirstPage: false });

  const chunks = [];
  pdf.on('data', d => chunks.push(d));
  const done = new Promise(resolve => pdf.on('end', () => resolve(Buffer.concat(chunks))));

  pdf.addPage();
  const page = pdf.page;

  // Área útil de la página
  const maxW = page.width  - page.margins.left - page.margins.right;
  const maxH = page.height - page.margins.top  - page.margins.bottom;

  // Escala manteniendo proporción
  let drawW = maxW, drawH = maxH;
  if (width && height) {
    const scale = Math.min(maxW / width, maxH / height);
    drawW = Math.floor(width  * scale);
    drawH = Math.floor(height * scale);
  }
  const x = page.margins.left + (maxW - drawW) / 2;
  const y = page.margins.top  + (maxH - drawH) / 2;

  pdf.image(pngBuffer, x, y, { width: drawW, height: drawH });
  pdf.end();
  return done;
}

/** Busca (o crea) subcarpeta YYYY-MM-DD bajo el folderId de la tienda. Devuelve su id. */
async function ensureDatedSubfolder(parentId, dateStr) {
  // Buscar por nombre exacto en ese parent
  const q = `name='${dateStr}' and mimeType='application/vnd.google-apps.folder' and '${parentId}' in parents and trashed=false`;
  const found = await withRetry('drive.list datedFolder', () =>
    driveApi.files.list({ q, fields: 'files(id,name)', spaces: 'drive', pageSize: 1 })
  );
  if (found.data.files?.length) return found.data.files[0].id;

  // Crear
  const folder = await withRetry('drive.create datedFolder', () =>
    driveApi.files.create({
      requestBody: {
        name: dateStr,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentId],
      },
      fields: 'id,name',
    })
  );
  return folder.data.id;
}

/** Sube un buffer PDF a Drive con nombre y parentId. */
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
 *  MAIN FLOW
 * ======================= */
async function main() {
  // 0) Token de acceso (para la descarga PNG)
  const accessToken = await auth.getAccessToken();
  if (!accessToken) {
    console.error('❌ No se pudo obtener accessToken del Service Account.');
    process.exit(1);
  }

  // 1) SheetId + charts por pestaña
  const byTitle = await getSheetsAndCharts();

  // 2) Por tienda → localizar sheetName, iterar charts
  const limit = pLimit(CONCURRENCY);
  let total = 0;

  for (const [tienda, cfg] of Object.entries(TIENDAS)) {
    const { sheetName, folderId } = cfg;

    const sh = byTitle.get(sheetName);
    if (!sh) {
      console.log(`⚠️ Hoja no encontrada: "${sheetName}" (${tienda}). Me la salto.`);
      continue;
    }
    if (!Array.isArray(sh.charts) || sh.charts.length === 0) {
      console.log(`ℹ️ ${tienda} / ${sheetName}: sin gráficos incrustados. Me la salto.`);
      continue;
    }

    // 2.1) Asegurar subcarpeta YYYY-MM-DD en Drive
    let dateFolderId;
    try {
      dateFolderId = await ensureDatedSubfolder(folderId, DATE_STR);
    } catch (e) {
      console.log(`❌ No puedo usar la carpeta de ${tienda} (${folderId}): ${e.message || e}. Me la salto.`);
      continue;
    }

    console.log(`🗂️ ${tienda} / ${sheetName}: ${sh.charts.length} gráficos → carpeta fecha ${DATE_STR}`);

    // 2.2) Exportar cada chart
    const tasks = sh.charts.map((c, i) => limit(async () => {
      const idx = i + 1;
      const title = (c.title || `${FILE_PREFIX}_${idx}`).replace(/[\\/:*?"<>|]/g, '_').slice(0, 80);
      const fileName = `${tienda}__${title}__${DATE_STR}.pdf`;
      try {
        const png = await downloadChartPNG({ sheetId: sh.sheetId, chartId: c.chartId, accessToken });
        const pdf = await pngToPDFBuffer(png, fileName);
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

main().catch(err => {
  console.error('💥 Error no controlado:', err);
  process.exit(1);
});
