export async function processData(dataBuffer, imageBuffer) {
  const workbook = XLSX.read(dataBuffer, { type: "array" });
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const products = mapWorksheetToProducts(worksheet);
  const groups = groupProducts(products);

  const newWorksheet = XLSX.utils.aoa_to_sheet([]);
  newWorksheet["!ref"] = XLSX.utils.encode_range({ r: 0, c: 0 }, { r: products.length * 15, c: 30 });

  const cellAddress = { r: 0, c: 1 };
  const idMap = drawGroups(cellAddress, groups, newWorksheet);

  const newWorkbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, "Report");

  const newBuffer = XLSX.write(newWorkbook, { type: "array", bookType: "xlsx" });
  const finalWorkbook = await processWorksheet(newBuffer, imageBuffer, idMap);

  return finalWorkbook;
}

class Product {
  category; // 0
  id; // 1
  name; // 2
  purchasePrice; // 3
  cost; // 4
  retail; // 5
  price; // 6
  avgPrice; // 9
  markup; // 10
  netSalesUnits; // 7
  finalBalance; // 12
  netSalesSum; // 8
  retailSoldPercentage; // 11
  currency; // 13
  material; // 14

  constructor(
    category,
    id,
    name,
    purchasePrice,
    cost,
    retail,
    price,
    avgPrice,
    markup,
    netSalesUnits,
    finalBalance,
    netSalesSum,
    retailSoldPercentage,
    currency,
    material
  ) {
    this.category = category;
    this.id = id;
    this.name = name;
    this.purchasePrice = purchasePrice;
    this.cost = cost;
    this.retail = retail;
    this.price = price;
    this.avgPrice = avgPrice.toFixed(2);
    this.markup = markup;
    this.netSalesUnits = netSalesUnits;
    this.finalBalance = finalBalance;
    this.netSalesSum = netSalesSum;
    this.retailSoldPercentage = retailSoldPercentage;
    this.currency = currency;
    this.material = material;
  }
}

class ProductGroup {
  category;
  products;

  constructor(category, products) {
    this.category = category;
    this.products = products;
  }
}

const mapWorksheetToProducts = (worksheet) => {
  const data = XLSX.utils.sheet_to_json(worksheet, { header: 1, range: 1, blankrows: false });

  let products = [];
  for (let i = 0, size = data.length; i < size; i++) {
    const obj = data[i];
    products.push(
      new Product(
        obj[0],
        obj[1],
        obj[2],
        obj[3],
        obj[4],
        obj[5],
        obj[6],
        obj[9],
        obj[10],
        obj[7],
        obj[12],
        obj[8],
        obj[11],
        obj[13],
        obj[14]
      )
    );
  }

  return products;
};

const groupProducts = (products) => {
  const productMap = products.reduce((acc, product) => {
    if (!acc[product.category]) {
      acc[product.category] = [];
    }

    acc[product.category].push(product);
    return acc;
  }, {});

  return Object.keys(productMap).map((key) => {
    return new ProductGroup(key, productMap[key]);
  });
};

const borders = {
  top: { style: "thin" },
  bottom: { style: "thin" },
  left: { style: "thin" },
  right: { style: "thin" },
};

const allBordersStyleCentered = {
  border: borders,
  alignment: { horizontal: "center" },
};

const emptyBorderedCell = { v: "", s: allBordersStyleCentered };

const drawGroups = (location, groups, worksheet) => {
  let baseRow = location.r;
  const map = new Map();

  for (let i = 0; i < groups.length; i += 1) {
    // Calculate the starting row for the current group
    const groupLocation = { r: baseRow, c: location.c };

    // Draw the current group
    drawGroup(groupLocation, groups[i], worksheet, map);

    // Calculate the number of rows the current group occupied
    // Each subgroup of 4 products takes 17 rows, and we add 2 rows for the header
    const groupRows = Math.ceil(groups[i].products.length / 3) * 17 + 2;

    // Update the baseRow for the next group, adding an additional row for extra spacing
    baseRow += groupRows + 1;
  }

  return map;
};

const drawGroup = (location, group, worksheet, map) => {
  worksheet[XLSX.utils.encode_cell(location)] = {
    v: group.category,
    s: {
      font: {
        bold: true,
        sz: "14",
      },
    },
  };

  for (let i = 0; i < group.products.length; i += 3) {
    let subGroup = group.products.slice(i, i + 3);
    let subGroupIndex = i / 3;

    for (let j = 0; j < subGroup.length; j++) {
      const loc = { r: location.r + 2 + subGroupIndex * 17, c: location.c + j * 5 };
      drawTable(loc, subGroup[j], worksheet);
      map.set(subGroup[j].id, { r: loc.r + 1, c: loc.c });
    }
  }
};

const drawTable = (location, product, worksheet) => {
  const headerLocation = { r: location.r + 10, c: location.c };
  drawHeader(location, product.name, worksheet);
  drawDataCells(headerLocation, product, worksheet);
};

const drawHeader = (location, productName, worksheet) => {
  const imageMerge = {
    s: location,
    e: { r: location.r + 9, c: location.c + 3 },
  };
  const headerLocation = { r: location.r + 10, c: location.c };
  const headerMerge = {
    s: headerLocation,
    e: { r: headerLocation.r + 1, c: headerLocation.c + 3 },
  };
  if (!worksheet["!merges"]) worksheet["!merges"] = [];

  worksheet["!merges"].push(headerMerge, imageMerge);

  worksheet[XLSX.utils.encode_cell(headerLocation)] = {
    v: productName,
    t: "s",
    s: allBordersStyleCentered,
  };

  worksheet[XLSX.utils.encode_cell({ r: headerLocation.r + 1, c: headerLocation.c })] = emptyBorderedCell;
  for (let i = 1; i < 4; i++) {
    worksheet[XLSX.utils.encode_cell({ r: headerLocation.r, c: headerLocation.c + i })] = emptyBorderedCell;
    worksheet[XLSX.utils.encode_cell({ r: headerLocation.r + 1, c: headerLocation.c + i })] = emptyBorderedCell;
  }
};

const drawDataCells = (headerLocation, product, worksheet) => {
  const rowIndex = headerLocation.r + 2;
  const column = headerLocation.c;
  const cell00 = cellAtIndex({ r: rowIndex, c: column });
  const cell01 = cellAtIndex({ r: rowIndex, c: column + 1 });
  const cell02 = cellAtIndex({ r: rowIndex, c: column + 2 });
  const cell03 = cellAtIndex({ r: rowIndex, c: column + 3 });
  const cell10 = cellAtIndex({ r: rowIndex + 1, c: column });
  const cell11 = cellAtIndex({ r: rowIndex + 1, c: column + 1 });
  const cell12 = cellAtIndex({ r: rowIndex + 1, c: column + 2 });
  const cell13 = cellAtIndex({ r: rowIndex + 1, c: column + 3 });
  const cell20 = cellAtIndex({ r: rowIndex + 2, c: column });
  const cell21 = cellAtIndex({ r: rowIndex + 2, c: column + 1 });
  const cell22 = cellAtIndex({ r: rowIndex + 2, c: column + 2 });
  const cell23 = cellAtIndex({ r: rowIndex + 2, c: column + 3 });
  const cell30 = cellAtIndex({ r: rowIndex + 3, c: column });
  const cell31 = cellAtIndex({ r: rowIndex + 3, c: column + 1 });
  const cell32 = cellAtIndex({ r: rowIndex + 3, c: column + 2 });
  const cell33 = cellAtIndex({ r: rowIndex + 3, c: column + 3 });

  worksheet[cell00] = cellWithValue(product.cost);
  worksheet[cell01] = cellWithValue(product.price);
  worksheet[cell02] = cellWithValue(product.avgPrice);
  worksheet[cell03] = cellWithValue(product.markup);

  worksheet[cell10] = cellWithValue(product.retail);
  worksheet[cell11] = cellWithValue(product.netSalesUnits);
  worksheet[cell12] = cellWithValue(product.finalBalance);
  worksheet[cell13] = emptyBorderedCell;

  worksheet[cell20] = cellWithValue(product.netSalesSum);

  const retailSoldPercentage = `${product.retailSoldPercentage}%`;
  const positiveFill = {
    fgColor: {
      rgb: "91e079",
    },
  };

  const negativeFill = {
    fgColor: {
      rgb: "e07979",
    },
  };

  if (product.retailSoldPercentage > 50) {
    worksheet[cell21] = cellWithValueAndStyle(retailSoldPercentage, {
      border: borders,
      fill: positiveFill,
      alignment: { horizontal: "center" },
    });
  } else if (product.retailSoldPercentage < 30) {
    worksheet[cell21] = cellWithValueAndStyle(retailSoldPercentage, {
      border: borders,
      fill: negativeFill,
      alignment: { horizontal: "center" },
    });
  } else {
    worksheet[cell21] = cellWithValue(retailSoldPercentage);
  }
  worksheet[cell22] = emptyBorderedCell;
  worksheet[cell23] = emptyBorderedCell;

  worksheet[cell30] = cellWithValue("Hind");
  worksheet[cell31] = cellWithValue(product.purchasePrice);
  worksheet[cell32] = cellWithValue(product.currency);
  worksheet[cell33] = cellWithValue(product.material);
};

const cellAtIndex = (address) => {
  return XLSX.utils.encode_cell(address);
};

const cellWithValue = (value) => {
  return cellWithValueAndStyle(value, allBordersStyleCentered);
};

const cellWithValueAndStyle = (value, style) => {
  return { v: value, s: style };
};

const processWorksheet = async (dataWorksheetBuffer, imageWorksheetBuffer, idMap) => {
  const dataWorkbook = await readWorkbookBuffer(dataWorksheetBuffer).catch((err) => console.error(err));

  await readImageWorkbookBuffer(imageWorksheetBuffer, dataWorkbook, idMap);
  return dataWorkbook;
};

async function readWorkbookBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);

  const worksheet = workbook.worksheets[0];

  const outWorkbook = new ExcelJS.Workbook();
  const outWorksheet = outWorkbook.addWorksheet("Report");

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

  return outWorkbook;
}

function getCellAtIndex(worksheet, r, c) {
  return worksheet.getRow(r + 1).getCell(c + 1);
}

function extractFirstNumber(str) {
  const result = str.match(/\d+/);
  return result ? parseInt(result[0], 10) : null;
}

async function readImageWorkbookBuffer(inputBuffer, inputWorkbook, idMap) {
  const imagesWorkbook = new ExcelJS.Workbook();
  await imagesWorkbook.xlsx.load(inputBuffer);
  const imagesWorksheet = imagesWorkbook.worksheets[0];
  const inputWorksheet = inputWorkbook.worksheets[0];

  for (const image of imagesWorksheet.getImages()) {
    const img = imagesWorkbook.model.media.find((m) => m.index === image.imageId);

    const imageRow = image.range.tl.nativeRow;
    const imageCol = image.range.tl.nativeCol;
    const textCellValue = getCellAtIndex(imagesWorksheet, imageRow + 1, imageCol).value;
    const id = extractFirstNumber(textCellValue);
    const address = idMap.get(id);

    if (img) {
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
