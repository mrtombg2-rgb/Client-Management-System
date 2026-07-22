/**
 * ======================================================
 * Products.gs
 * Product Management Module
 * ======================================================
 */

const PRODUCT_SHEET = "Products";

/**
 * Generate Product ID
 * Example: PR000001
 */
function generateProductId_() {

  const sh = SpreadsheetApp.getActive().getSheetByName(PRODUCT_SHEET);

  if (!sh || sh.getLastRow() < 2) {
    return "PR000001";
  }

  const ids = sh.getRange(2,1,sh.getLastRow()-1,1)
                .getValues()
                .flat();

  let max = 0;

  ids.forEach(id => {

    if (!id) return;

    const num = parseInt(
      String(id).replace("PR",""),
      10
    );

    if (!isNaN(num) && num > max) {
      max = num;
    }

  });

  return "PR" + Utilities.formatString("%06d", max + 1);

}

/**
 * Add Product
 */
function addProduct(name, usd, bdt) {

  const sh = SpreadsheetApp.getActive().getSheetByName(PRODUCT_SHEET);

  const id = generateProductId_();

  sh.appendRow([
    id,
    name,
    Number(usd),
    Number(bdt),
    "Active",
    new Date()
  ]);

  return id;

}

/**
 * Get Product By Name
 */
function getProductByName(name){

  const sh = SpreadsheetApp.getActive().getSheetByName(PRODUCT_SHEET);

  if(sh.getLastRow()<2) return null;

  const data = sh.getRange(2,1,sh.getLastRow()-1,6).getValues();

  for(let i=0;i<data.length;i++){

    if(String(data[i][1]).trim()==String(name).trim()){

      return{

        row:i+2,
        id:data[i][0],
        name:data[i][1],
        usd:data[i][2],
        bdt:data[i][3],
        status:data[i][4],
        updated:data[i][5]

      };

    }

  }

  return null;

}

/**
 * Get Product By ID
 */
function getProductById(id){

  const sh = SpreadsheetApp.getActive().getSheetByName(PRODUCT_SHEET);

  if(sh.getLastRow()<2) return null;

  const data = sh.getRange(2,1,sh.getLastRow()-1,6).getValues();

  for(let i=0;i<data.length;i++){

    if(data[i][0]==id){

      return{

        row:i+2,
        id:data[i][0],
        name:data[i][1],
        usd:data[i][2],
        bdt:data[i][3],
        status:data[i][4],
        updated:data[i][5]

      };

    }

  }

  return null;

}

/**
 * Update Product Price
 */
function updateProductPrice(productId, usd, bdt){

  const product = getProductById(productId);

  if(!product) return false;

  const sh = SpreadsheetApp.getActive().getSheetByName(PRODUCT_SHEET);

  sh.getRange(product.row,3).setValue(Number(usd));
  sh.getRange(product.row,4).setValue(Number(bdt));
  sh.getRange(product.row,6).setValue(new Date());

  return true;

}

/**
 * Activate Product
 */
function activateProduct(productId){

  const product = getProductById(productId);

  if(!product) return false;

  SpreadsheetApp.getActive()
    .getSheetByName(PRODUCT_SHEET)
    .getRange(product.row,5)
    .setValue("Active");

  return true;

}

/**
 * Deactivate Product
 */
function deactivateProduct(productId){

  const product = getProductById(productId);

  if(!product) return false;

  SpreadsheetApp.getActive()
    .getSheetByName(PRODUCT_SHEET)
    .getRange(product.row,5)
    .setValue("Inactive");

  return true;

}

/**
 * Get Active Products
 */
function getActiveProducts(){

  const sh = SpreadsheetApp.getActive().getSheetByName(PRODUCT_SHEET);

  if(sh.getLastRow()<2) return [];

  const data = sh.getRange(2,1,sh.getLastRow()-1,6).getValues();

  return data.filter(r => String(r[4])==="Active");

}

/**
 * Product Price Lookup
 * Used by Transactions.gs
 */
function getProductPrice(productName){

  const product = getProductByName(productName);

  if(!product) return null;

  return {

    usd: Number(product.usd),
    bdt: Number(product.bdt)

  };

}

/**
 * Product Dropdown List
 */
function getProductList(){

  return getActiveProducts().map(function(r){

    return r[1];

  });

}

/**
 * Check Product Exists
 */
function productExists(productName){

  return getProductByName(productName)!=null;

}
