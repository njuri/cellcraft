const ExcelJS = require("exceljs");

const processWorkseet = async (dataWorksheetBuffer, imageWorksheet, idMaps) => {
  const dataWorkbook = await readWorkbookBuffer(dataWorksheetBuffer).catch((err) => console.error(err));
  console.log(idMaps.length);
  await readImageWorkbook(imageWorksheet, dataWorkbook, idMaps);

  dataWorkbook.xlsx.writeFile("output/out_with_images.xlsx");
};

async function readWorkbookBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const outWorkbook = new ExcelJS.Workbook();

  for (const worksheet of workbook.worksheets) {
    const outWorksheet = outWorkbook.addWorksheet(worksheet.name);

    // Copy cells and styles
    worksheet.eachRow({ includeEmpty: true }, function (row, rowNumber) {
      row.eachCell({ includeEmpty: true }, function (cell, colNumber) {
        const newCell = outWorksheet.getRow(rowNumber).getCell(colNumber);
        newCell.value = cell.value;
        newCell.style = cell.style;
      });
    });

    // Copy merged cells
    worksheet.model.merges.forEach((merge) => {
      outWorksheet.mergeCells(merge);
    });
  }

  return outWorkbook;
}

function getCellAtIndex(worksheet, r, c) {
  return worksheet.getRow(r + 1).getCell(c + 1);
}

function extractFirstNumber(str) {
  const result = str.match(/\d+/);
  return result ? parseInt(result[0], 10) : null;
}

async function readImageWorkbook(filePath, inputWorkbook, idMaps) {
  const imagesWorkbook = new ExcelJS.Workbook();
  await imagesWorkbook.xlsx.readFile(filePath);
  const imagesWorksheet = imagesWorkbook.worksheets[0];

  for (const [i, inputWorksheet] of inputWorkbook.worksheets.entries()) {
    for (const image of imagesWorksheet.getImages()) {
      const img = imagesWorkbook.model.media.find((m) => m.index === image.imageId);

      const imageRow = image.range.tl.nativeRow;
      const imageCol = image.range.tl.nativeCol;
      const textCellValue = getCellAtIndex(imagesWorksheet, imageRow + 1, imageCol).value;
      const id = extractFirstNumber(textCellValue);
      const address = idMaps[i].get(id);

      if (img && address) {
        const imageId = inputWorkbook.addImage({
          buffer: img.buffer,
          extension: img.extension,
        });

        inputWorksheet.addImage(imageId, {
          tl: { col: address.c, row: address.r - 1 },
          br: { col: address.c + 4, row: address.r + 9 },
        });
      }
    }
  }
}

module.exports = { processWorkseet };
