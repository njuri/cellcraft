export async function processData(dataBuffer) {
  const workbook = await readWorkbookBuffer(dataBuffer).catch((err) => console.error(err));

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

  return outWorkbook;
}

async function readWorkbookBuffer(buffer) {
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(buffer);
  const outWorkbook = new ExcelJS.Workbook();

  for (const worksheet of workbook.worksheets) {
    const outWorksheet = outWorkbook.addWorksheet(worksheet.name);

    // Copy cells and styles
    worksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
      row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const newCell = outWorksheet.getRow(rowNumber).getCell(colNumber);
        newCell.value = cell.value;
        newCell.style = cell.style;
      });
    });

    // Copy merged cells
    for (const merge of worksheet.model.merges) {
      outWorksheet.mergeCells(merge);
    }
  }

  return outWorkbook;
}

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
  season; // 15
  manufacturer; // 16

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
    material,
    season,
    manufacturer,
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
    this.season = season;
    this.manufacturer = manufacturer;
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

class ManufacturerProducts {
  manufacturer;
  groups;

  constructor(manufacturer, groups) {
    this.manufacturer = manufacturer;
    this.groups = groups;
  }

  productCount() {
    return this.groups.reduce((totalCount, group) => {
      return totalCount + group.products.length;
    }, 0);
  }
}

class SeasonProducts {
  season;
  groups;

  constructor(season, groups) {
    this.season = season;
    this.groups = groups;
  }

  productCount() {
    return this.groups.reduce((totalCount, group) => {
      return totalCount + group.products.length;
    }, 0);
  }
}

class OrderProductSection {
  name;
  groups;

  constructor(name, groups) {
    this.name = name;
    this.groups = groups;
  }

  artTotalEE() {
    return this.groups.reduce((sum, group) => {
      if (typeof group.artEE === "function") {
        return sum + group.artEE();
      }
      return sum;
    }, 0);
  }

  artTotalLV() {
    return this.groups.reduce((sum, group) => {
      if (typeof group.artLV === "function") {
        return sum + group.artLV();
      }
      return sum;
    }, 0);
  }

  artTotalLT() {
    return this.groups.reduce((sum, group) => {
      if (typeof group.artLT === "function") {
        return sum + group.artLT();
      }
      return sum;
    }, 0);
  }

  orderTotalEE() {
    return this.groups.reduce((sum, group) => {
      if (typeof group.sumEE === "function") {
        return sum + group.sumEE();
      }
      return sum;
    }, 0);
  }

  orderTotalLV() {
    return this.groups.reduce((sum, group) => {
      if (typeof group.sumLV === "function") {
        return sum + group.sumLV();
      }
      return sum;
    }, 0);
  }

  orderTotalLT() {
    return this.groups.reduce((sum, group) => {
      if (typeof group.sumLT === "function") {
        return sum + group.sumLT();
      }
      return sum;
    }, 0);
  }
}

class OrderProduct {
  manufacturer;
  currency; // f
  category; // j
  sex; // k
  inBoxAmount; // v
  price; // x
  amountEE; // r
  amountLV; // s
  amountLT; // t

  constructor(
    manufacturer,
    currency,
    category,
    sex,
    inBoxAmount,
    price,
    amountEE,
    amountLV,
    amountLT,
  ) {
    this.manufacturer = manufacturer;
    this.category = category;
    this.currency = currency;
    this.sex = sex;
    this.inBoxAmount = inBoxAmount;
    this.price = price;
    this.amountEE = amountEE;
    this.amountLV = amountLV;
    this.amountLT = amountLT;
  }

  artEE() {
    return this.amountEE > 0 ? 1 : 0;
  }

  artLV() {
    return this.amountLV > 0 ? 1 : 0;
  }

  artLT() {
    return this.amountLT > 0 ? 1 : 0;
  }

  priceEur() {
    const usdCurrency = "USD";
    const rmbCurrency = "RMB";

    function usdToEur(usd) {
      return usd / 1;
    }

    function rmbToEur(rmb) {
      return rmb / 7.5;
    }

    if (this.currency === usdCurrency) {
      return usdToEur(this.price);
    }
    if (this.currency === rmbCurrency) {
      return rmbToEur(this.price);
    }
  }

  sumEE() {
    return this.sumEur(this.amountEE);
  }

  sumLV() {
    return this.sumEur(this.amountLV);
  }

  sumLT() {
    return this.sumEur(this.amountLT);
  }

  sumEur(amount) {
    const effectiveAmount = amount ?? 0;
    return this.inBoxAmount * effectiveAmount * this.priceEur();
  }
}

ProductGroup.prototype.artEE = function () {
  return this.products.reduce((sum, product) => {
    if (typeof product.artEE === "function") {
      return sum + product.artEE();
    }
    return sum;
  }, 0);
};

ProductGroup.prototype.artLV = function () {
  return this.products.reduce((sum, product) => {
    if (typeof product.artLV === "function") {
      return sum + product.artLV();
    }
    return sum;
  }, 0);
};

ProductGroup.prototype.artLT = function () {
  return this.products.reduce((sum, product) => {
    if (typeof product.artLT === "function") {
      return sum + product.artLT();
    }
    return sum;
  }, 0);
};

ProductGroup.prototype.sumEE = function () {
  return this.products.reduce((sum, product) => {
    if (typeof product.sumEE === "function") {
      return sum + product.sumEE();
    }
    return sum;
  }, 0);
};

ProductGroup.prototype.sumLV = function () {
  return this.products.reduce((sum, product) => {
    if (typeof product.sumLV === "function") {
      return sum + product.sumLV();
    }
    return sum;
  }, 0);
};

ProductGroup.prototype.sumLT = function () {
  return this.products.reduce((sum, product) => {
    if (typeof product.sumLT === "function") {
      return sum + product.sumLT();
    }
    return sum;
  }, 0);
};

const groupOrderProductsBySection = (products) => {
  const productMap = products.reduce((acc, product) => {
    const section = product.sex.toLowerCase();

    if (!acc[section]) {
      acc[section] = [];
    }

    acc[section].push(product);
    return acc;
  }, {});

  return Object.keys(productMap).map((key) => {
    return new OrderProductSection(key, groupProductsByCategoryCaseInsensitive(productMap[key]));
  });
};

const groupProductsByCategoryCaseInsensitive = (products) => {
  const productMap = products.reduce((acc, product) => {
    const category = product.category.toLowerCase();

    if (!acc[category]) {
      acc[category] = [];
    }

    acc[category].push(product);
    return acc;
  }, {});

  return Object.keys(productMap).map((key) => {
    return new ProductGroup(key, productMap[key]);
  });
};
