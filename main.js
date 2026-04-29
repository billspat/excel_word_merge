
const { app, BrowserWindow, dialog, ipcMain, shell } = require('electron');
const path = require('path');
const fs = require('fs');
const ExcelJS = require('exceljs');
const PizZip = require('pizzip');
const Docxtemplater = require('docxtemplater');

let win;
let rows = [];
let excelBaseName = 'merged';

function createWindow() {
  win = new BrowserWindow({
    width: 1200,
    height: 800,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js')
    }
  });
  win.loadFile('index.html');
}

app.whenReady().then(createWindow);

// Extract every {placeholder} name from a .docx template.
// Strips XML tags first so run-split placeholders are still found.
function getTemplatePlaceholders(templatePath) {
  const content = fs.readFileSync(templatePath, 'binary');
  const zip = new PizZip(content);
  const candidates = [
    'word/document.xml', 'word/header1.xml', 'word/header2.xml',
    'word/header3.xml', 'word/footer1.xml', 'word/footer2.xml', 'word/footer3.xml'
  ];
  let fullText = '';
  for (const f of candidates) {
    if (zip.files[f]) fullText += zip.files[f].asText();
  }
  const text = fullText.replace(/<[^>]+>/g, '');
  const tags = new Set();
  const re = /\{([^{}\n\r\t]+)\}/g;
  let m;
  while ((m = re.exec(text)) !== null) tags.add(m[1].trim());
  return tags;
}

ipcMain.handle('open-excel', async () => {
  const { filePaths } = await dialog.showOpenDialog({ filters: [{ extensions: ['xlsx'] }] });
  if (!filePaths || filePaths.length === 0) return null;
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePaths[0]);
  const sheet = workbook.worksheets[0];
  const headers = sheet.getRow(1).values.slice(1).map(v => String(v));
  rows = [];
  sheet.eachRow((row, rowNum) => {
    if (rowNum === 1) return;
    const obj = {};
    headers.forEach((h, i) => { obj[h] = row.getCell(i + 1).value; });
    rows.push(obj);
  });
  excelBaseName = path.basename(filePaths[0], '.xlsx');
  return { headers, rows, fileName: path.basename(filePaths[0]) };
});

ipcMain.handle('validate-schema', async (event, templatePath) => {
  const issues = [];

  // Check 1: Excel must be tabular — headers + at least one data row
  if (!rows || rows.length === 0) {
    issues.push('No data found. Load an Excel file that has column headings and at least one data row.');
    return { ok: false, issues };
  }
  const headers = Object.keys(rows[0]);
  if (headers.length === 0) {
    issues.push('The Excel file has no column headings in the first row.');
    return { ok: false, issues };
  }

  // Check 2 & 3: Compare headers against template placeholders
  if (!templatePath) {
    issues.push('No Word template selected. Select a template before validating.');
    return { ok: false, issues };
  }
  const placeholders = getTemplatePlaceholders(templatePath);

  const headersNotInTemplate = headers.filter(h => !placeholders.has(h));
  const placeholdersNotInExcel = [...placeholders].filter(p => !headers.includes(p));

  if (headersNotInTemplate.length > 0)
    issues.push(`Column headings with no matching template placeholder: ${headersNotInTemplate.join(', ')}`);
  if (placeholdersNotInExcel.length > 0)
    issues.push(`Template placeholders with no matching column heading: ${placeholdersNotInExcel.join(', ')}`);

  return { ok: issues.length === 0, issues };
});

ipcMain.handle('select-template', async () => {
  const { filePaths } = await dialog.showOpenDialog({ filters: [{ extensions: ['docx'] }] });
  return filePaths && filePaths.length > 0 ? filePaths[0] : null;
});

ipcMain.handle('merge-docs', async (event, templatePath) => {
  const { filePaths: folderPaths, canceled } = await dialog.showOpenDialog({
    title: 'Select output folder',
    properties: ['openDirectory', 'createDirectory']
  });
  if (canceled || !folderPaths || folderPaths.length === 0) return { canceled: true };

  const outDir = path.join(folderPaths[0], excelBaseName);
  fs.mkdirSync(outDir, { recursive: true });

  for (let i = 0; i < rows.length; i++) {
    const content = fs.readFileSync(templatePath, 'binary');
    const zip = new PizZip(content);
    const doc = new Docxtemplater(zip, { paragraphLoop: true, linebreaks: true });
    doc.render(rows[i]);
    const fname = `${rows[i].Semester}_${rows[i].AssistantNetID} review.docx`;
    fs.writeFileSync(path.join(outDir, fname), doc.getZip().generate({ type: 'nodebuffer' }));
    win.webContents.send('progress', Math.round(((i + 1) / rows.length) * 100));
  }
  shell.openPath(outDir);
  return { canceled: false, outDir };
});
