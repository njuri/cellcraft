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

module.exports = {
  Product,
  ProductGroup,
  ManufacturerProducts,
  SeasonProducts,
};
