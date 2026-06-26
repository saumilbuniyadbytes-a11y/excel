const fs = require("fs");
const ExcelJS = require("exceljs");
const path = require("path");

// Create workbook with ExcelJS directly (no need for temp file)
async function createExcelWithImages(json_file) {
  const data = JSON.parse(fs.readFileSync(json_file, "utf8"));
  const workbook = new ExcelJS.Workbook();
  const sheet = workbook.addWorksheet("BBS");

  // Get headers from first object
  const headers = Object.keys(data[0]);

  // Add header row with styling
  const headerRow = sheet.addRow(headers);
  headerRow.eachCell((cell) => {
    cell.font = { bold: true };
    cell.fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFD9E1F2" },
    };
    cell.border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
    cell.alignment = { horizontal: "center", vertical: "middle" };
  });

  // Find Shape column index (1-based for ExcelJS)
  const shapeColIndex = headers.indexOf("Shape") + 1;

  // Add data rows
  for (let i = 0; i < data.length; i++) {
    const row = data[i];
    const values = headers.map((h) => (row[h] === null ? "" : row[h]));
    const excelRow = sheet.addRow(values);

    // Add borders to all cells
    excelRow.eachCell((cell) => {
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });

    // Handle image
    let imagePath = row["Shape"];

    // Skip if no valid path
    if (
      !imagePath ||
      imagePath === null ||
      imagePath === "null" ||
      imagePath === ""
    ) {
      // console.log(`Row ${i + 2}: No image path`);
      continue;
    }

    // Convert to string and clean
    imagePath = String(imagePath).trim();

    // Try adding .png if no extension
    if (!path.extname(imagePath)) {
      imagePath = imagePath + ".png";
    }

    // Check if file exists
    if (!fs.existsSync(imagePath)) {
      // console.log(`Row ${i + 2}: Image not found - ${imagePath}`);
      continue;
    }

    // Get extension
    const ext = path.extname(imagePath).slice(1).toLowerCase();

    // Add image to workbook
    const imageId = workbook.addImage({
      filename: imagePath,
      extension: ext,
    });

    // Set row height for image
    excelRow.height = 80;

    // Insert image into cell
    sheet.addImage(imageId, {
      tl: { col: shapeColIndex - 1, row: i + 1 }, // 0-based
      ext: { width: 100, height: 80 },
    });

    // Clear the cell text
    excelRow.getCell(shapeColIndex).value = "";

    console.log(`Row ${i + 2}: Image added - ${imagePath}`);
  }

  // Auto-fit column widths
  sheet.columns.forEach((column) => {
    let maxLength = 0;
    column.eachCell({ includeEmpty: true }, (cell) => {
      const columnLength = cell.value ? String(cell.value).length : 10;
      if (columnLength > maxLength) {
        maxLength = columnLength;
      }
    });
    column.width = Math.min(maxLength + 2, 30);
  });

  // Set Shape column width specifically for image
  if (shapeColIndex > 0) {
    sheet.getColumn(shapeColIndex).width = 15;
  }

  // Save file
  await workbook.xlsx.writeFile("final.xlsx");
  console.timeEnd("total");
  console.log("Excel file created: final.xlsx");
}
console.time("total");

const json_file = "output.json";

createExcelWithImages(json_file);
