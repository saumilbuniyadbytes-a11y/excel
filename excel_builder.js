// console.time("total");
const XLSX = require("xlsx");
const fs = require("fs");
const ExcelJS = require("exceljs");
const path = require("path");

async function makeExcel(data) {
  // console.time("building time");
  // make new file
  const workbook = XLSX.utils.book_new();
  // make new worksheet
  const worksheet = XLSX.utils.json_to_sheet(data);
  // add sheet in file
  XLSX.utils.book_append_sheet(workbook, worksheet, "Beam");
  XLSX.writeFile(workbook, "final.xlsx");
  // console.timeEnd("building time");
}

async function image(data) {
  console.time("total");
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile("final.xlsx");
  const worksheet = workbook.getWorksheet("Beam");
  const headerRow = worksheet.getRow(1);
  let shapeCol;

  headerRow.eachCell((cell, colNumber) => {
    if (cell.value === "Shape") {
      shapeCol = colNumber;
    }
  });

  data.forEach((item, index) => {
    let imagePath = item.Shape;

    if (!imagePath) {
      // console.log(`Row ${index + 2}: No image`);
      return;
    }

    imagePath = String(imagePath).trim();

    if (!path.extname(imagePath)) {
      imagePath += ".png";
    }

    if (!fs.existsSync(imagePath)) {
      // console.log(`Row ${index + 2}: image not found - "${imagePath}"`);
      return;
    }

    const imageId = workbook.addImage({
      filename: imagePath,
      extension: path.extname(imagePath).substring(1),
    });

    const excelRow = index + 2;

    worksheet.addImage(imageId, {
      tl: { col: shapeCol - 1, row: excelRow - 1 },
      ext: { width: 60, height: 60 },
    });
    worksheet.getRow(excelRow).height = 50;
  });
  await workbook.xlsx.writeFile("final.xlsx");

  console.timeEnd("total");
}

async function main(json_file) {
  const data = JSON.parse(fs.readFileSync(json_file, "utf-8"));
  await makeExcel(data);
  await image(data);
}
const json_file = "output.json";

// createExcelWithImages(json_file).catch(console.error);
main(json_file);
// console.timeEnd("total");
