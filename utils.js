const XLSX = require("xlsx-js-style");
const {
  Product,
  ProductGroup,
  ManufacturerProducts,
  SeasonProducts,
  OrderProductSection,
} = require("./Product.js");

const mapWorksheetToProducts = (worksheet) => {
  const data = XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    range: 1,
    blankrows: false,
  });

  const products = [];
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
        obj[14],
        obj[15],
        obj[16],
      ),
    );
  }

  return products;
};

const groupProductsByCategory = (products) => {
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

const groupProductsByManufacturer = (products) => {
  const productMap = products.reduce((acc, product) => {
    const manufacturerLower = product.manufacturer.toLowerCase();

    if (!acc[manufacturerLower]) {
      acc[manufacturerLower] = [];
    }

    acc[manufacturerLower].push(product);
    return acc;
  }, {});

  return Object.keys(productMap).map((key) => {
    return new ManufacturerProducts(
      key,
      groupProductsByCategory(productMap[key]),
    );
  });
};

const groupProductsBySeason = (products) => {
  const productMap = products.reduce((acc, product) => {
    const seasonLower = product.season.toLowerCase();

    if (!acc[seasonLower]) {
      acc[seasonLower] = [];
    }

    acc[seasonLower].push(product);
    return acc;
  }, {});

  return Object.keys(productMap).map((key) => {
    return new SeasonProducts(key, groupProductsByCategory(productMap[key]));
  });
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
    return new OrderProductSection(
      key,
      groupProductsByCategory(productMap[key]),
    );
  });
};

module.exports = {
  mapWorksheetToProducts,
  groupProductsByCategory,
  groupProductsByManufacturer,
  groupProductsBySeason,
  groupOrderProductsBySection,
};
