const ExcelJS = require("exceljs");

function extractFirstNumber(str) {
  const result = str.match(/\d+/);
  return result ? parseInt(result[0], 10) : null;
}

function getCellAtIndex(worksheet, r, c) {
  return worksheet.getRow(r + 1).getCell(c + 1);
}

async function readWorkbook(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const worksheet = workbook.worksheets[0];

  const newWorkbook = new ExcelJS.Workbook();
  const newWorksheet = newWorkbook.addWorksheet("My Sheet");
  newWorksheet.getCell("A1").value = "Hello world!";

  for (const image of worksheet.getImages()) {
    const img = workbook.model.media.find((m) => m.index === image.imageId);

    const imageRow = image.range.tl.nativeRow;
    const imageCol = image.range.tl.nativeCol;
    const textCellValue = getCellAtIndex(worksheet, imageRow + 1, imageCol).value;
    const id = extractFirstNumber(textCellValue);
    console.log(id);

    if (img) {
      const imageId = newWorkbook.addImage({
        buffer: img.buffer,
        extension: img.extension,
      });

      const row = newWorksheet.getRow(image.range.tl.nativeRow + 1);
      row.height = 160;
      const column = newWorksheet.getColumn(image.range.tl.nativeCol + 1);
      column.width = 37;

      newWorksheet.addImage(imageId, {
        tl: { col: image.range.tl.nativeCol, row: image.range.tl.nativeRow },
        br: { col: image.range.tl.nativeCol + 1, row: image.range.tl.nativeRow + 1 },
      });
    }
  }

  await newWorkbook.xlsx.writeFile("output/ExcelJS.xlsx");
  console.log("Workbook created and saved successfully!");
}

readWorkbook("res/KKKPx.xlsx").catch((err) => console.error(err));

module.exports = { drawImages };
