const XLSX = require("xlsx-js-style");
const ExcelJS = require("exceljs");
const fs = require("fs");

const {
  mapWorksheetToProducts,
  groupProductsByCategory,
  groupProductsByManufacturer,
  groupProductsBySeason,
  groupOrderProductsBySection,
} = require("./utils");
const { OrderProduct } = require("./Product.js");

async function readExcelFile(filePath) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.readFile(filePath);

  const startSuffix = " START";
  const endSuffix = " END";
  const products = [];

  workbook.eachSheet((worksheet, id) => {
    let manufacturer = "";
    let shouldStopReading = false;
    worksheet.eachRow((row, rowNumber) => {
      if (!shouldStopReading) {
        const headerCell = row.values[1];
        if (headerCell?.endsWith(startSuffix)) {
          manufacturer = headerCell.slice(0, -startSuffix.length);
        } else if (headerCell?.endsWith(endSuffix)) {
          shouldStopReading = true;
        }
        if (manufacturer) {
          const product = new OrderProduct(
            manufacturer,
            row.values[6],
            row.values[10],
            row.values[11],
            row.values[22],
            row.values[24],
            row.values[18],
            row.values[19],
            row.values[20],
          );
          products.push(product);
        }
      }
    });
  });

  const outWorkbook = new ExcelJS.Workbook();
  const outWorksheet = outWorkbook.addWorksheet("Result");

  const sections = groupOrderProductsBySection(products);

  drawTable({ r: 5, c: 1 }, sections, outWorksheet);

  outWorkbook.xlsx.writeFile("output/result.xlsx");
}

readExcelFile("res/input.xlsx")
  .then(() => {
    console.log("Reading completed.");
  })
  .catch((error) => {
    console.error("Error reading the Excel file:", error);
  });

// const { drawGroups } = require("./drawing");
// const { processWorkseet } = require("./image-processor");

// const workbook = XLSX.readFile("res/s_DATA.xlsx");
// const sheetName = workbook.SheetNames[0];
// const worksheet = workbook.Sheets[sheetName];
// const products = mapWorksheetToProducts(worksheet);
// const allGroups = groupProductsByCategory(products);

// const newWorksheet = XLSX.utils.aoa_to_sheet([]);
// newWorksheet["!ref"] = XLSX.utils.encode_range(
//   { r: 0, c: 0 },
//   { r: products.length * 20, c: 16 },
// );
// const cellAddress = { r: 0, c: 1 };
// const headerMaps = [];
// headerMaps.push(drawGroups(cellAddress, allGroups, newWorksheet));
// const newWorkbook = XLSX.utils.book_new();
// XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Report");

// const groupedByManufacturer = groupProductsByManufacturer(products);
// for (const group of groupedByManufacturer) {
//   const manufacturerWorksheet = XLSX.utils.aoa_to_sheet([]);

//   manufacturerWorksheet["!ref"] = XLSX.utils.encode_range(
//     { r: 0, c: 0 },
//     { r: group.productCount() * 20, c: 16 },
//   );
//   const cellAddress = { r: 0, c: 1 };

//   headerMaps.push(drawGroups(cellAddress, group.groups, manufacturerWorksheet));
//   XLSX.utils.book_append_sheet(
//     newWorkbook,
//     manufacturerWorksheet,
//     group.manufacturer.substring(0, 30),
//   );
// }

// const groupedBySeason = groupProductsBySeason(products);
// for (const group of groupedBySeason) {
//   const seasonWorksheet = XLSX.utils.aoa_to_sheet([]);

//   seasonWorksheet["!ref"] = XLSX.utils.encode_range(
//     { r: 0, c: 0 },
//     { r: group.productCount() * 20, c: 16 },
//   );
//   const cellAddress = { r: 0, c: 1 };

//   headerMaps.push(drawGroups(cellAddress, group.groups, seasonWorksheet));
//   XLSX.utils.book_append_sheet(
//     newWorkbook,
//     seasonWorksheet,
//     group.season.substring(0, 30).replace(/[\\\/\?\*\[\]]/g, ""),
//   );
// }

// XLSX.writeFile(newWorkbook, "output/out_no_images.xlsx");

// const dataBuf = XLSX.write(newWorkbook, { type: "buffer", bookType: "xlsx" });

// processWorkseet(dataBuf, "res/s_FOTO.xlsx", headerMaps);

const drawTable = (location, sections, worksheet) => {
  let offset = 0;
  drawHeadings(location, worksheet);

  for (const [index, section] of sections.entries()) {
    const sectionLocation = { r: location.r + offset + index, c: location.c };
    drawHeader(sectionLocation, worksheet);
    offset += section.groups.length + 4;
    for (const [index, group] of section.groups.entries()) {
      drawOrderGroup({ r: sectionLocation.r + 2 + index, c: sectionLocation.c }, group, worksheet);
    }
    drawTotalRows(
      {
        r: sectionLocation.r + section.groups.length + 2,
        c: sectionLocation.c,
      },
      section.name,
      worksheet,
      section.orderTotalEE(),
      section.artTotalEE(),
      section.orderTotalLV(),
      section.artTotalLV(),
      section.orderTotalLT(),
      section.artTotalLT(),
    );
  }
};

const drawHeadings = (location, worksheet) => {
  worksheet.getCell(location.r, location.c + 1).value = "EE";
  worksheet.getCell(location.r, location.c + 2).value = "EE";
  worksheet.getCell(location.r, location.c + 3).value = "LV";
  worksheet.getCell(location.r, location.c + 4).value = "LV";
  worksheet.getCell(location.r, location.c + 5).value = "LT";
  worksheet.getCell(location.r, location.c + 6).value = "LT";
};

const boldTextStyle = { font: { bold: true } };
const allBordersStyle = {
  border: {
    top: { style: "thin" },
    left: { style: "thin" },
    bottom: { style: "thin" },
    right: { style: "thin" },
  },
};

const yellowFill = {
  fill: {
    type: "pattern",
    pattern: "solid",
    fgColor: { argb: "FFFF00" },
  },
};

const drawHeader = (location, worksheet) => {
  const art = "Арт в AW24";
  const order = "AW24 Заказ";

  worksheet.getCell(location.r + 1, location.c).value = "Вид обуви";
  worksheet.getCell(location.r + 1, location.c).style = boldTextStyle;
  worksheet.getCell(location.r + 1, location.c).border = allBordersStyle;

  worksheet.getCell(location.r + 1, location.c + 1).value = order;
  worksheet.getCell(location.r + 1, location.c + 1).style = {
    ...boldTextStyle,
    ...allBordersStyle,
  };

  worksheet.getCell(location.r + 1, location.c + 2).value = art;
  worksheet.getCell(location.r + 1, location.c + 2).style = boldTextStyle;
  worksheet.getCell(location.r + 1, location.c + 2).border = allBordersStyle;

  worksheet.getCell(location.r + 1, location.c + 3).value = order;
  worksheet.getCell(location.r + 1, location.c + 3).style = boldTextStyle;
  worksheet.getCell(location.r + 1, location.c + 3).border = allBordersStyle;

  worksheet.getCell(location.r + 1, location.c + 4).value = art;
  worksheet.getCell(location.r + 1, location.c + 4).style = boldTextStyle;
  worksheet.getCell(location.r + 1, location.c + 4).border = allBordersStyle;

  worksheet.getCell(location.r + 1, location.c + 5).value = order;
  worksheet.getCell(location.r + 1, location.c + 5).style = boldTextStyle;
  worksheet.getCell(location.r + 1, location.c + 5).border = allBordersStyle;

  worksheet.getCell(location.r + 1, location.c + 6).value = art;
  worksheet.getCell(location.r + 1, location.c + 6).style = boldTextStyle;
  worksheet.getCell(location.r + 1, location.c + 6).border = allBordersStyle;
};

const drawOrderGroup = (location, group, worksheet) => {
  worksheet.getCell(location.r, location.c).value = group.category;
  const config = new CellConfig(group.category, { ...allBordersStyle, ...boldTextStyle });
  updateCell(location.r, location.c, worksheet, config);

  worksheet.getCell(location.r, location.c + 1).value = group.sumEE();
  worksheet.getCell(location.r, location.c + 2).value = group.artEE();
  worksheet.getCell(location.r, location.c + 3).value = group.sumLV();
  worksheet.getCell(location.r, location.c + 4).value = group.artLV();
  worksheet.getCell(location.r, location.c + 5).value = group.sumLT();
  worksheet.getCell(location.r, location.c + 6).value = group.artLT();
};

const drawTotalRows = (
  location,
  name,
  worksheet,
  orderTotalEE,
  artTotalEE,
  orderTotalLV,
  artTotalLV,
  orderTotalLT,
  artTotalLT,
) => {
  worksheet.getCell(location.r, location.c).value = "Oсень";
  worksheet.getCell(location.r + 1, location.c).value = "Зима";

  worksheet.getCell(location.r + 2, location.c).value = `Итого ${name}`;
  const config = new CellConfig(`Итого ${name}`, {
    ...boldTextStyle,
    ...yellowFill,
  });
  updateCell(location.r + 2, location.c, worksheet, config);

  const config2 = new CellConfig(orderTotalEE, {
    ...boldTextStyle,
    ...yellowFill,
  });
  updateCell(location.r + 2, location.c + 1, worksheet, config2);

  worksheet.getCell(location.r + 2, location.c + 2).value = artTotalEE;
  worksheet.getCell(location.r + 2, location.c + 2).style = boldTextStyle;

  worksheet.getCell(location.r + 2, location.c + 3).value = orderTotalLV;
  worksheet.getCell(location.r + 2, location.c + 3).style = boldTextStyle;

  worksheet.getCell(location.r + 2, location.c + 4).value = artTotalLV;
  worksheet.getCell(location.r + 2, location.c + 4).style = boldTextStyle;

  worksheet.getCell(location.r + 2, location.c + 5).value = orderTotalLT;
  worksheet.getCell(location.r + 2, location.c + 5).style = boldTextStyle;

  worksheet.getCell(location.r + 2, location.c + 6).value = artTotalLT;
  worksheet.getCell(location.r + 2, location.c + 6).style = boldTextStyle;
};

class CellConfig {
  value;
  style;

  constructor(value, style) {
    this.value = value;
    this.style = style;
  }
}

const updateCell = (r, c, worksheet, config) => {
  worksheet.getCell(r, c).value = config.value;

  worksheet.getCell(r, c).style = {
    ...(worksheet.getCell(r, c).style || {}),
    ...config.style,
  };
};
