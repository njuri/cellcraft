const XLSX = require("xlsx-js-style");
const { Product, ProductGroup } = require("./Product.js");

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

module.exports = { mapWorksheetToProducts, groupProducts };
