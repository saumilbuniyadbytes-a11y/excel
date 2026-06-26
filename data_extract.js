const XLSX = require('xlsx');
const fs = require('fs');
const ExcelJS = require('exceljs');
const AdmZip = require('adm-zip');
const path = require('path');

async function main() {
  const workbook = XLSX.readFile('Tricity-Tricity.xlsx');
  const sheet = workbook.Sheets[workbook.SheetNames[0]];

  const rows = XLSX.utils.sheet_to_json(sheet, {
    header: 1
  });

  const headers = rows[4];

  const data = rows
    .slice(5)
    .filter(row => row.length)
    .map(row => {
      const obj = {};
      headers.forEach((header, index) => {
        obj[String(header).trim()] = row[index] ?? null;
      });
      return obj;
    });

  // =========================
  // Extract images from XLSX
  // =========================
  const imageDir = './images';
  fs.mkdirSync(imageDir, { recursive: true });

  const zip = new AdmZip('Tricity-Tricity.xlsx');
  const extractedImages = [];

  zip.getEntries()
    .filter(entry => entry.entryName.startsWith('xl/media/'))
    .forEach((entry, index) => {
      const ext = path.extname(entry.entryName);
      const fileName = `shape_${index + 1}${ext}`;

      fs.writeFileSync(
        path.join(imageDir, fileName),
        entry.getData()
      );

      extractedImages[index] = fileName;
    });

  // =========================
  // Read image positions
  // =========================
  const excelWorkbook = new ExcelJS.Workbook();
  await excelWorkbook.xlsx.readFile('Tricity-Tricity.xlsx');

  const worksheet = excelWorkbook.getWorksheet(1);
  const rowImageMap = new Map();

  worksheet.getImages().forEach(img => {
    const excelRow = img.range.tl.row + 1; 
    const fileName = extractedImages[img.imageId] || null;

    // Directly assign the path string (e.g., "images/shape_1.png")
    const imagePath = fileName ? path.join(imageDir, fileName) : null;

    rowImageMap.set(excelRow, imagePath);
  });

  // =========================
  // Attach image path to data
  // =========================
  data.forEach((item, index) => {
    const excelRow = index + 6; // First data row

    item.Shape = rowImageMap.get(excelRow) || null;
  });

  // =========================
  // Save JSON
  // =========================
  fs.writeFileSync(
    'output.json',
    JSON.stringify(data, null, 2)
  );

  console.log('JSON saved');
  console.log('Images saved to ./images');
}

main().catch(console.error);