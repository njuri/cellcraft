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

module.exports = {
  Product,
  ProductGroup,
  ManufacturerProducts,
  SeasonProducts,
  OrderProduct,
  OrderProductSection,
};
