const XLSX = require("xlsx-js-style");
const ExcelJS = require("exceljs");
const fs = require("fs");

const { mapWorksheetToProducts, groupProducts } = require("./utils");
const { drawGroups } = require("./drawing");

const workbook = XLSX.readFile("res/kk_table.xlsx");
const sheetName = workbook.SheetNames[0];
const worksheet = workbook.Sheets[sheetName];
const products = mapWorksheetToProducts(worksheet);
const groups = groupProducts(products);
const newWorksheet = XLSX.utils.aoa_to_sheet([]);
newWorksheet["!ref"] = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: products.length * 100, c: 30 });

const cellAddress = { r: 0, c: 1 };
drawGroups(cellAddress, groups, newWorksheet);

const newWorkbook = XLSX.utils.book_new();
XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Report");
XLSX.writeFile(newWorkbook, "output/newfile.xlsx");
