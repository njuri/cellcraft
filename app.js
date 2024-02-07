const XLSX = require("xlsx-js-style");
const ExcelJS = require("exceljs");
const fs = require("fs");

const {
  mapWorksheetToProducts,
  groupProductsByCategory,
  groupProductsByManufacturer,
  groupProductsBySeason,
} = require("./utils");
const { drawGroups } = require("./drawing");
const { processWorkseet } = require("./image-processor");

const workbook = XLSX.readFile("res/s_DATA.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const products = mapWorksheetToProducts(worksheet);
const allGroups = groupProductsByCategory(products);

const newWorksheet = XLSX.utils.aoa_to_sheet([]);
newWorksheet["!ref"] = XLSX.utils.encode_range(
  { r: 0, c: 0 },
  { r: products.length * 20, c: 16 },
);
const cellAddress = { r: 0, c: 1 };
const headerMaps = [];
headerMaps.push(drawGroups(cellAddress, allGroups, newWorksheet));
const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Report");

const groupedByManufacturer = groupProductsByManufacturer(products);
for (const group of groupedByManufacturer) {
  const manufacturerWorksheet = XLSX.utils.aoa_to_sheet([]);

  manufacturerWorksheet["!ref"] = XLSX.utils.encode_range(
    { r: 0, c: 0 },
    { r: group.productCount() * 20, c: 16 },
  );
  const cellAddress = { r: 0, c: 1 };

  headerMaps.push(drawGroups(cellAddress, group.groups, manufacturerWorksheet));
  XLSX.utils.book_append_sheet(
    newWorkbook,
    manufacturerWorksheet,
    group.manufacturer.substring(0, 30),
  );
}

const groupedBySeason = groupProductsBySeason(products);
for (const group of groupedBySeason) {
  const seasonWorksheet = XLSX.utils.aoa_to_sheet([]);

  seasonWorksheet["!ref"] = XLSX.utils.encode_range(
    { r: 0, c: 0 },
    { r: group.productCount() * 20, c: 16 },
  );
  const cellAddress = { r: 0, c: 1 };

  headerMaps.push(drawGroups(cellAddress, group.groups, seasonWorksheet));
  XLSX.utils.book_append_sheet(
    newWorkbook,
    seasonWorksheet,
    group.season.substring(0, 30).replace(/[\\\/\?\*\[\]]/g, ""),
  );
}

XLSX.writeFile(newWorkbook, "output/out_no_images.xlsx");

const dataBuf = XLSX.write(newWorkbook, { type: "buffer", bookType: "xlsx" });

processWorkseet(dataBuf, "res/s_FOTO.xlsx", headerMaps);
