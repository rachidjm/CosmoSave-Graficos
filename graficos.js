const { google } = require('googleapis');

async function crearGraficosParaTodas({ auth, spreadsheetId }) {
  const sheets = google.sheets({ version: 'v4', auth });
  const ss = await sheets.spreadsheets.get({ spreadsheetId });

  const hojas = ss.data.sheets
    .map(s => ({
      id: s.properties.sheetId,
      title: s.properties.title,
      type: s.properties.sheetType || 'GRID'
    }))
    // 👇 Solo hojas normales (GRID) y que no sean auxiliares
    .filter(s =>
      s.type === 'GRID' &&
      s.title &&
      !s.title.endsWith(' - Calc') &&
      !s.title.endsWith(' - Gráfico')
    );

  for (const h of hojas) {
    await crearGraficosDeTienda({ auth, spreadsheetId, tienda: h.title, sheetId: h.id });
  }
}

async function crearGraficosDeTienda({ auth, spreadsheetId, tienda, sheetId }) {
  const sheets = google.sheets({ version: 'v4', auth });

  // 0) Asegurar columna I = PrecioNum (convierte "52,99 €" a 52.99)
  const formula =
    '=ARRAYFORMULA(IF(ROW(F:F)=1,"PrecioNum",IF(F2:F="","",IFERROR(VALUE(REGEXREPLACE(SUBSTITUTE(SUBSTITUTE(F2:F,".",""),",","."),"[^0-9.-]",""))))))';

  const reqs = [
    {
      updateCells: {
        range: { sheetId, startRowIndex: 0, endRowIndex: 1, startColumnIndex: 8, endColumnIndex: 9 },
        rows: [{ values: [{ userEnteredValue: { stringValue: 'PrecioNum' } }] }],
        fields: 'userEnteredValue'
      }
    },
    {
      updateCells: {
        range: { sheetId, startRowIndex: 1, endRowIndex: 2, startColumnIndex: 8, endColumnIndex: 9 },
        rows: [{ values: [{ userEnteredValue: { formulaValue: formula } }] }],
        fields: 'userEnteredValue'
      }
    }
  ];

  // 1) Crear/limpiar hoja de cálculos (pivots)
  const ss = await sheets.spreadsheets.get({ spreadsheetId });
  const calcTitle = `${tienda} - Calc`;
  const calc = ss.data.sheets.find(s => s.properties.title === calcTitle);
  let calcId;
  if (calc) {
    calcId = calc.properties.sheetId;
    // limpiar contenido (deja la hoja vacía)
    reqs.push({
      updateCells: {
        range: { sheetId: calcId, startRowIndex: 0, startColumnIndex: 0 },
        fields: 'userEnteredValue'
      }
    });
  } else {
    const add = await sheets.spreadsheets.batchUpdate({
      spreadsheetId,
      requestBody: { requests: [{ addSheet: { properties: { title: calcTitle } } }] }
    });
    calcId = add.data.replies[0].addSheet.properties.sheetId;
  }

  // helpers pivot + charts
  const pivot = ({ rows, values }) => ({
    updateCells: {
      range: { sheetId: calcId, startRowIndex: 0, startColumnIndex: 0, endRowIndex: 100000, endColumnIndex: 26 },
      rows: [{
        values: [{
          pivotTable: {
            source: { sheetId, startRowIndex: 1, startColumnIndex: 0 }, // A2:...
            // FIX: obliga orden en cada grupo (evita "No sort order specified")
            rows: (rows || []).map(r => ({
              sortOrder: 'ASCENDING',
              showTotals: false,
              ...r
            })),
            values,
            valueLayout: 'HORIZONTAL'
          }
        }]
      }],
      fields: 'pivotTable'
    }
  });

  const lineChart = ({ title, xCol, yCol }) => ({
    addChart: {
      chart: {
        spec: {
          title,
          basicChart: {
            chartType: 'LINE',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Fecha' },
              { position: 'LEFT_AXIS', title: 'Valor' }
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: calcId, startRowIndex: 1, startColumnIndex: xCol, endColumnIndex: xCol + 1 }]
                }
              }
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{ sheetId: calcId, startRowIndex: 1, startColumnIndex: yCol, endColumnIndex: yCol + 1 }]
                }
              },
              targetAxis: 'LEFT_AXIS'
            }]
          }
        },
        position: { newSheet: true }
      }
    }
  });

  const pieChart = ({ title, labelCol, valueCol }) => ({
    addChart: {
      chart: {
        spec: {
          title,
          pieChart: {
            legendPosition: 'RIGHT_LEGEND',
            domain: { sourceRange: { sources: [{ sheetId: calcId, startRowIndex: 1, startColumnIndex: labelCol, endColumnIndex: labelCol + 1 }] } },
            series: { sourceRange: { sources: [{ sheetId: calcId, startRowIndex: 1, startColumnIndex: valueCol, endColumnIndex: valueCol + 1 }] } }
          }
        },
        position: { newSheet: true }
      }
    }
  });

  const barChart = ({ title, xCol, yCol }) => ({
    addChart: {
      chart: {
        spec: {
          title,
          basicChart: {
            chartType: 'BAR',
            legendPosition: 'BOTTOM_LEGEND',
            axis: [
              { position: 'BOTTOM_AXIS', title: 'Clicks' },
              { position: 'LEFT_AXIS', title: 'Elemento' }
            ],
            domains: [{
              domain: {
                sourceRange: {
                  sources: [{ sheetId: calcId, startRowIndex: 1, startColumnIndex: xCol, endColumnIndex: xCol + 1 }]
                }
              }
            }],
            series: [{
              series: {
                sourceRange: {
                  sources: [{ sheetId: calcId, startRowIndex: 1, startColumnIndex: yCol, endColumnIndex: yCol + 1 }]
                }
              },
              targetAxis: 'LEFT_AXIS'
            }]
          }
        },
        position: { newSheet: true }
      }
    }
  });

  // G1) Clicks por día (Fecha vs COUNT Producto) → LINE
  reqs.push(pivot({
    rows: [{ sourceColumnOffset: 0 }],             // A = Fecha
    values: [{ summarizeFunction: 'COUNTA', sourceColumnOffset: 2 }] // C = Producto
  }));
  reqs.push(lineChart({ title: `Clicks por día — ${tienda}`, xCol: 0, yCol: 1 }));

  // G2) Dispositivo (E) → PIE
  reqs.push(pivot({
    rows: [{ sourceColumnOffset: 4 }],             // E = Dispositivo
    values: [{ summarizeFunction: 'COUNTA', sourceColumnOffset: 4 }]
  }));
  reqs.push(pieChart({ title: `Dispositivo — ${tienda}`, labelCol: 0, valueCol: 1 }));

  // G3) Top productos (C) → BAR
  reqs.push(pivot({
    rows: [{ sourceColumnOffset: 2 }],             // C = Producto
    values: [{ summarizeFunction: 'COUNTA', sourceColumnOffset: 2 }]
  }));
  reqs.push(barChart({ title: `Top productos — ${tienda}`, xCol: 1, yCol: 0 }));

  // G4) Precio medio diario (A vs AVG I) → LINE
  reqs.push(pivot({
    rows: [{ sourceColumnOffset: 0 }],             // A = Fecha
    values: [{ summarizeFunction: 'AVERAGE', sourceColumnOffset: 8 }] // I = PrecioNum
  }));
  reqs.push(lineChart({ title: `Precio medio por día — ${tienda}`, xCol: 0, yCol: 1 }));

  await sheets.spreadsheets.batchUpdate({ spreadsheetId, requestBody: { requests: reqs } });
}

module.exports = { crearGraficosParaTodas, crearGraficosDeTienda };
