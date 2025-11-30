/**
 * Inventory Management System for Google Sheets
 * Main Code.gs file with all server-side functions
 */

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Inventory App')
    .addItem('Add Sale', 'openAddSaleSidebar')
    .addItem('Edit Sale', 'openEditSaleSidebar')
    .addItem('Delete Sale', 'openDeleteSaleSidebar')
    .addSeparator()
    .addItem('Generate Monthly Report', 'generateMonthlyReport')
    .addItem('Sync Sales View to Data', 'syncSalesViewToData')
    .addSeparator()
    .addItem('Add Product', 'openAddProductSidebar')
    .addItem('Edit Product', 'openEditProductSidebar')
    .addItem('Delete Product', 'openDeleteProductSidebar')
    // .addSeparator()
    // .addItem('Update Product Names in Sales', 'updateSaleItemsWithProductNames')
    .addToUi();
}

// =========================
// PRODUCT FUNCTIONS
// =========================

/**
 * Get all products from Products sheet
 */
function getProducts() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Products');
  if (!sheet) {
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) {
    return [];
  }

  // Get headers to find correct column indices
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const findColIndex = (possibleNames, defaultIndex) => {
    for (const name of possibleNames) {
      const index = headers.indexOf(name);
      if (index >= 0) return index;
    }
    return defaultIndex; // Fallback to default position
  };

  const idCol = findColIndex(['ID'], 0);
  const nameCol = findColIndex(['Name'], 1);
  const descCol = findColIndex(['Description', 'Desc'], 2);
  const priceCol = findColIndex(['Price', 'Selling Price'], 3);
  const discountCol = findColIndex(['Discount', 'Discount Price'], 4);
  const stockCol = findColIndex(['Stock'], 5);

  const rows = sheet
    .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
    .getValues();
  const products = rows.map((r) => ({
    id: r[idCol],
    name: r[nameCol],
    desc: r[descCol],
    price: r[priceCol],
    discount: r[discountCol],
    stock: r[stockCol],
  }));
  // Sort by name A-Z
  products.sort((a, b) => (a.name || '').localeCompare(b.name || ''));
  return products;
}

/**
 * Open sidebar to add a new product
 */
function openAddProductSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('form')
    .setTitle('Add Product')
    .setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Submit a new product to the Products sheet
 */
function submitProduct(name, desc, price, discount, stock) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Products');
    if (!sheet) {
      throw new Error(
        "Products sheet not found. Please create a sheet named 'Products' with headers: ID, Name, Description, Price, Discount, Stock, Created At",
      );
    }

    // Get current headers to check structure
    let headers = [];
    if (sheet.getLastRow() > 0) {
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    }

    const hasLastUpdated = headers.includes('Last Updated At');
    const hasCreatedAt = headers.includes('Created At');

    // Create sheet with headers if it doesn't exist
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'ID',
        'Name',
        'Description',
        'Price',
        'Discount',
        'Stock',
        'Created At',
        'Last Updated At',
      ]);
      headers = [
        'ID',
        'Name',
        'Description',
        'Price',
        'Discount',
        'Stock',
        'Created At',
        'Last Updated At',
      ];

      // Format date/time columns in header row
      const createdAtHeaderCol = headers.indexOf('Created At') + 1;
      const lastUpdatedHeaderCol = headers.indexOf('Last Updated At') + 1;
      if (createdAtHeaderCol > 0) {
        sheet
          .getRange(1, createdAtHeaderCol)
          .setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }
      if (lastUpdatedHeaderCol > 0) {
        sheet
          .getRange(1, lastUpdatedHeaderCol)
          .setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }
    } else if (!hasLastUpdated) {
      // Insert Last Updated At column after Created At (column 7, index 6)
      const createdAtIndex = headers.indexOf('Created At');
      if (createdAtIndex >= 0) {
        sheet.insertColumnAfter(createdAtIndex + 1);
        sheet.getRange(1, createdAtIndex + 2).setValue('Last Updated At');
        // Refresh headers after insertion
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

        // Format the new Last Updated At column for all existing rows
        const lastUpdatedColIndex = headers.indexOf('Last Updated At') + 1;
        if (lastUpdatedColIndex > 0 && sheet.getLastRow() > 1) {
          sheet
            .getRange(2, lastUpdatedColIndex, sheet.getLastRow() - 1, 1)
            .setNumberFormat('yyyy-mm-dd hh:mm:ss');
        }
      } else {
        // Created At not found, add at the end
        sheet.insertColumnAfter(sheet.getLastColumn());
        sheet.getRange(1, sheet.getLastColumn()).setValue('Last Updated At');
        headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

        // Format the new Last Updated At column for all existing rows
        const lastUpdatedColIndex = sheet.getLastColumn();
        if (sheet.getLastRow() > 1) {
          sheet
            .getRange(2, lastUpdatedColIndex, sheet.getLastRow() - 1, 1)
            .setNumberFormat('yyyy-mm-dd hh:mm:ss');
        }
      }

      // Also format Created At column if it exists
      const createdColIndex = headers.indexOf('Created At') + 1;
      if (createdColIndex > 0 && sheet.getLastRow() > 1) {
        sheet
          .getRange(2, createdColIndex, sheet.getLastRow() - 1, 1)
          .setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }
    }

    const id = 'PRD-' + new Date().getTime();
    const createdAt = new Date();
    const lastUpdated = new Date();

    // Find column indices (handle different header name variations)
    const findColumnIndex = (possibleNames, defaultIndex) => {
      for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index >= 0) return index + 1;
      }
      return defaultIndex + 1; // Fallback to default position
    };

    const idCol = findColumnIndex(['ID'], 0);
    const nameCol = findColumnIndex(['Name'], 1);
    const descCol = findColumnIndex(['Description', 'Desc'], 2);
    const priceCol = findColumnIndex(['Price', 'Selling Price'], 3);
    const discountCol = findColumnIndex(['Discount', 'Discount Price'], 4);
    const stockCol = findColumnIndex(['Stock'], 5);
    const createdCol = findColumnIndex(['Created At', 'Created'], 6);
    const lastUpdatedCol = findColumnIndex(
      ['Last Updated At', 'Last Updated'],
      7,
    );

    // Validate column indices
    if (
      idCol <= 0 ||
      nameCol <= 0 ||
      descCol <= 0 ||
      priceCol <= 0 ||
      discountCol <= 0 ||
      stockCol <= 0 ||
      createdCol <= 0
    ) {
      throw new Error(
        'Required columns not found in Products sheet. Found headers: ' +
          headers.join(', '),
      );
    }

    // Get the next row number
    const nextRow = sheet.getLastRow() + 1;

    // Set values in correct columns
    sheet.getRange(nextRow, idCol).setValue(id);
    sheet.getRange(nextRow, nameCol).setValue(name);
    sheet.getRange(nextRow, descCol).setValue(desc);
    sheet.getRange(nextRow, priceCol).setValue(Number(price));
    sheet.getRange(nextRow, discountCol).setValue(Number(discount || 0));
    sheet.getRange(nextRow, stockCol).setValue(Number(stock));

    // Set date/time values with proper formatting
    sheet.getRange(nextRow, createdCol).setValue(createdAt);
    sheet.getRange(nextRow, createdCol).setNumberFormat('yyyy-mm-dd hh:mm:ss');

    if (lastUpdatedCol > 0) {
      sheet.getRange(nextRow, lastUpdatedCol).setValue(lastUpdated);
      sheet
        .getRange(nextRow, lastUpdatedCol)
        .setNumberFormat('yyyy-mm-dd hh:mm:ss');
    }

    return { success: true, message: 'Product added successfully!' };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Open sidebar to edit a product
 */
function openEditProductSidebar() {
  const template = HtmlService.createTemplateFromFile('editProduct');
  const html = template.evaluate().setTitle('Edit Product').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Get product details by ID
 */
function getProductDetails(id) {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Products');
  if (!sheet) {
    return null;
  }

  // Get headers to find correct column indices
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const findColIndex = (possibleNames, defaultIndex) => {
    for (const name of possibleNames) {
      const index = headers.indexOf(name);
      if (index >= 0) return index;
    }
    return defaultIndex;
  };

  const idCol = findColIndex(['ID'], 0);
  const nameCol = findColIndex(['Name'], 1);
  const descCol = findColIndex(['Description', 'Desc'], 2);
  const priceCol = findColIndex(['Price', 'Selling Price'], 3);
  const discountCol = findColIndex(['Discount', 'Discount Price'], 4);
  const stockCol = findColIndex(['Stock'], 5);

  const rows = sheet.getDataRange().getValues();
  for (let i = 1; i < rows.length; i++) {
    if (rows[i][idCol] === id) {
      return {
        name: rows[i][nameCol],
        desc: rows[i][descCol],
        price: rows[i][priceCol],
        discount: rows[i][discountCol],
        stock: rows[i][stockCol],
      };
    }
  }
  return null;
}

/**
 * Update an existing product
 */
function updateProduct(id, name, desc, price, discount, stock) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Products');
    if (!sheet) {
      throw new Error('Products sheet not found');
    }

    // Check if Last Updated At column exists, add if not
    let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const hasLastUpdated = headers.includes('Last Updated At');
    if (!hasLastUpdated) {
      sheet.insertColumnAfter(sheet.getLastColumn());
      sheet.getRange(1, sheet.getLastColumn()).setValue('Last Updated At');
      // Refresh headers after insertion
      headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

      // Format the new Last Updated At column for all existing rows
      const lastUpdatedColIndex = sheet.getLastColumn();
      if (sheet.getLastRow() > 1) {
        sheet
          .getRange(2, lastUpdatedColIndex, sheet.getLastRow() - 1, 1)
          .setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }

      // Also format Created At column if it exists
      const createdColIndex = headers.indexOf('Created At') + 1;
      if (createdColIndex > 0 && sheet.getLastRow() > 1) {
        sheet
          .getRange(2, createdColIndex, sheet.getLastRow() - 1, 1)
          .setNumberFormat('yyyy-mm-dd hh:mm:ss');
      }
    }

    // Find column indices (handle different header name variations)
    const findColIndex = (possibleNames) => {
      for (const name of possibleNames) {
        const index = headers.indexOf(name);
        if (index >= 0) return index + 1;
      }
      return -1;
    };

    const nameCol = findColIndex(['Name']);
    const descCol = findColIndex(['Description', 'Desc']);
    const priceCol = findColIndex(['Price', 'Selling Price']);
    const discountCol = findColIndex(['Discount', 'Discount Price']);
    const stockCol = findColIndex(['Stock']);
    const lastUpdatedCol = findColIndex(['Last Updated At', 'Last Updated']);

    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        // Update product fields using correct column indices
        if (nameCol > 0) sheet.getRange(i + 1, nameCol).setValue(name);
        if (descCol > 0) sheet.getRange(i + 1, descCol).setValue(desc);
        if (priceCol > 0)
          sheet.getRange(i + 1, priceCol).setValue(Number(price));
        if (discountCol > 0)
          sheet.getRange(i + 1, discountCol).setValue(Number(discount || 0));
        if (stockCol > 0)
          sheet.getRange(i + 1, stockCol).setValue(Number(stock));

        // Update Last Updated At timestamp
        if (lastUpdatedCol > 0) {
          const updateTime = new Date();
          sheet.getRange(i + 1, lastUpdatedCol).setValue(updateTime);
          sheet
            .getRange(i + 1, lastUpdatedCol)
            .setNumberFormat('yyyy-mm-dd hh:mm:ss');
        }

        return { success: true, message: 'Product updated successfully!' };
      }
    }
    throw new Error('Product not found');
  } catch (error) {
    return { success: false, message: error.message };
  }
}

/**
 * Open sidebar to delete a product
 */
function openDeleteProductSidebar() {
  const template = HtmlService.createTemplateFromFile('deleteProduct');
  const html = template.evaluate().setTitle('Delete Product').setWidth(400);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Delete a product by ID
 */
function deleteProduct(id) {
  try {
    const sheet = SpreadsheetApp.getActive().getSheetByName('Products');
    if (!sheet) {
      throw new Error('Products sheet not found');
    }
    const rows = sheet.getDataRange().getValues();
    for (let i = 1; i < rows.length; i++) {
      if (rows[i][0] === id) {
        sheet.deleteRow(i + 1);
        return { success: true, message: 'Product deleted successfully!' };
      }
    }
    throw new Error('Product not found');
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =========================
// SALES FUNCTIONS
// =========================

/**
 * Open sidebar to add a new sale
 */
function openAddSaleSidebar() {
  const template = HtmlService.createTemplateFromFile('addSale');
  template.products = getProducts();
  const html = template.evaluate().setTitle('Add Sale').setWidth(550);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Submit a sale with multiple items
 */
function submitSale(saleItems, deliveryFee, saleDiscount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    deliveryFee = deliveryFee || 0;
    saleDiscount = saleDiscount || 0;

    // Get or create sheets
    let saleSheet = ss.getSheetByName('Sales');
    if (!saleSheet) {
      saleSheet = ss.insertSheet('Sales');
      saleSheet.appendRow([
        'Sale ID',
        'Date',
        'Subtotal',
        'Delivery Fee',
        'Discount',
        'Total Amount',
      ]);
    } else if (saleSheet.getLastRow() === 0) {
      saleSheet.appendRow([
        'Sale ID',
        'Date',
        'Subtotal',
        'Delivery Fee',
        'Discount',
        'Total Amount',
      ]);
    }

    // Check if Sales sheet needs new columns
    const saleHeaders = saleSheet
      .getRange(1, 1, 1, saleSheet.getLastColumn())
      .getValues()[0];
    if (!saleHeaders.includes('Subtotal')) {
      // Old format, need to add columns
      const totalColIndex = saleHeaders.indexOf('Total Amount');
      if (totalColIndex >= 0) {
        // Insert Subtotal, Delivery Fee, Discount before Total Amount
        saleSheet.insertColumnBefore(totalColIndex + 1);
        saleSheet.insertColumnBefore(totalColIndex + 1);
        saleSheet.insertColumnBefore(totalColIndex + 1);
        saleSheet.getRange(1, totalColIndex + 1).setValue('Subtotal');
        saleSheet.getRange(1, totalColIndex + 2).setValue('Delivery Fee');
        saleSheet.getRange(1, totalColIndex + 3).setValue('Discount');
      }
    }

    let itemSheet = ss.getSheetByName('SaleItems');
    if (!itemSheet) {
      itemSheet = ss.insertSheet('SaleItems');
      itemSheet.appendRow([
        'Sale ID',
        'Product ID',
        'Product Name',
        'Quantity',
        'Is Free',
        'Unit Price',
        'Final Price',
      ]);
    } else if (itemSheet.getLastRow() === 0) {
      itemSheet.appendRow([
        'Sale ID',
        'Product ID',
        'Product Name',
        'Quantity',
        'Is Free',
        'Unit Price',
        'Final Price',
      ]);
    }

    // Check if we need to add Product Name column to existing sheet
    const headers = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductName = headers.includes('Product Name');
    if (!hasProductName && itemSheet.getLastRow() > 0) {
      // Insert Product Name column after Product ID (column B, so insert at column C)
      itemSheet.insertColumnAfter(2);
      itemSheet.getRange(1, 3).setValue('Product Name');
    }

    const prodSheet = ss.getSheetByName('Products');
    if (!prodSheet) {
      throw new Error('Products sheet not found');
    }

    const saleId = 'SALE-' + new Date().getTime();
    const date = new Date();
    let totalAmount = 0;

    // Process each sale item
    saleItems.forEach((item) => {
      const products = getProducts();
      const prod = products.find((p) => p.id === item.productId);
      if (!prod) {
        throw new Error(`Product with ID ${item.productId} not found`);
      }

      // Check stock availability
      if (item.qty > prod.stock) {
        throw new Error(
          `Not enough stock for ${prod.name}. Available: ${prod.stock}, Requested: ${item.qty}`,
        );
      }

      // Calculate price
      // Use provided unit price if available, otherwise use product's discount or price
      let unitPrice;
      if (item.isFree) {
        unitPrice = 0;
      } else if (
        item.unitPrice !== null &&
        item.unitPrice !== undefined &&
        item.unitPrice !== ''
      ) {
        // Use the provided unit price
        unitPrice = Number(item.unitPrice);
      } else {
        // Use the default formula: discount if available, otherwise price
        unitPrice = prod.discount > 0 ? prod.discount : prod.price;
      }
      const finalPrice = unitPrice * item.qty;
      totalAmount += finalPrice;

      // Deduct stock and update Last Updated At
      const prodHeaders = prodSheet
        .getRange(1, 1, 1, prodSheet.getLastColumn())
        .getValues()[0];
      const hasLastUpdated = prodHeaders.includes('Last Updated At');

      // Add Last Updated At column if it doesn't exist
      let lastUpdatedColIndex = 0;
      if (!hasLastUpdated) {
        prodSheet.insertColumnAfter(prodSheet.getLastColumn());
        prodSheet
          .getRange(1, prodSheet.getLastColumn())
          .setValue('Last Updated At');
        lastUpdatedColIndex = prodSheet.getLastColumn();
        // Refresh headers after insertion
        prodHeaders = prodSheet
          .getRange(1, 1, 1, prodSheet.getLastColumn())
          .getValues()[0];

        // Format the new Last Updated At column for all existing rows
        if (prodSheet.getLastRow() > 1) {
          prodSheet
            .getRange(2, lastUpdatedColIndex, prodSheet.getLastRow() - 1, 1)
            .setNumberFormat('yyyy-mm-dd hh:mm:ss');
        }

        // Also format Created At column if it exists
        const createdColIndex = prodHeaders.indexOf('Created At') + 1;
        if (createdColIndex > 0 && prodSheet.getLastRow() > 1) {
          prodSheet
            .getRange(2, createdColIndex, prodSheet.getLastRow() - 1, 1)
            .setNumberFormat('yyyy-mm-dd hh:mm:ss');
        }
      } else {
        lastUpdatedColIndex = prodHeaders.indexOf('Last Updated At') + 1;
      }

      // Find column indices
      const findColIndex = (possibleNames, defaultIndex) => {
        for (const name of possibleNames) {
          const index = prodHeaders.indexOf(name);
          if (index >= 0) return index;
        }
        return defaultIndex;
      };

      const idCol = findColIndex(['ID'], 0);
      const stockCol = findColIndex(['Stock'], 5);

      const rows = prodSheet.getDataRange().getValues();
      for (let i = 1; i < rows.length; i++) {
        if (rows[i][idCol] === item.productId) {
          const newStock = rows[i][stockCol] - item.qty;
          prodSheet.getRange(i + 1, stockCol + 1).setValue(newStock);

          // Update Last Updated At timestamp
          if (lastUpdatedColIndex > 0) {
            const updateTime = new Date();
            prodSheet.getRange(i + 1, lastUpdatedColIndex).setValue(updateTime);
            prodSheet
              .getRange(i + 1, lastUpdatedColIndex)
              .setNumberFormat('yyyy-mm-dd hh:mm:ss');
          }
          break;
        }
      }

      // Insert line item with product name
      const productName = prod.name || 'Unknown Product';
      itemSheet.appendRow([
        saleId,
        item.productId,
        productName,
        item.qty,
        item.isFree,
        unitPrice,
        finalPrice,
      ]);
    });

    // Calculate final total
    const subtotal = totalAmount;
    const finalTotal = subtotal - deliveryFee - saleDiscount;

    // Insert sale summary with new columns
    // Find column indices for Sales sheet
    const updatedSaleHeaders = saleSheet
      .getRange(1, 1, 1, saleSheet.getLastColumn())
      .getValues()[0];
    const subtotalCol = updatedSaleHeaders.indexOf('Subtotal') + 1;
    const deliveryFeeCol = updatedSaleHeaders.indexOf('Delivery Fee') + 1;
    const discountCol = updatedSaleHeaders.indexOf('Discount') + 1;
    const totalCol = updatedSaleHeaders.indexOf('Total Amount') + 1;

    const nextSaleRow = saleSheet.getLastRow() + 1;
    saleSheet.getRange(nextSaleRow, 1).setValue(saleId);
    saleSheet.getRange(nextSaleRow, 2).setValue(date);
    if (subtotalCol > 0)
      saleSheet.getRange(nextSaleRow, subtotalCol).setValue(subtotal);
    if (deliveryFeeCol > 0)
      saleSheet.getRange(nextSaleRow, deliveryFeeCol).setValue(deliveryFee);
    if (discountCol > 0)
      saleSheet.getRange(nextSaleRow, discountCol).setValue(saleDiscount);
    if (totalCol > 0) {
      saleSheet.getRange(nextSaleRow, totalCol).setValue(finalTotal);
    } else {
      // Fallback for old format
      saleSheet.getRange(nextSaleRow, 3).setValue(finalTotal);
    }

    // Track delivery fee in DeliveryCosts sheet if there's a delivery fee
    if (deliveryFee > 0) {
      let deliverySheet = ss.getSheetByName('DeliveryCosts');
      if (!deliverySheet) {
        deliverySheet = ss.insertSheet('DeliveryCosts');
        deliverySheet.appendRow(['Sale ID', 'Date', 'Delivery Fee']);
      } else if (deliverySheet.getLastRow() === 0) {
        deliverySheet.appendRow(['Sale ID', 'Date', 'Delivery Fee']);
      }
      deliverySheet.appendRow([saleId, date, deliveryFee]);
    }

    // Automatically append new sale to Sales View sheet if it exists
    try {
      const viewSheet = ss.getSheetByName('Sales View');
      if (viewSheet && viewSheet.getLastRow() > 0) {
        // Only update if the sheet exists and has data
        // Use a small delay to ensure the data is written before reading
        Utilities.sleep(100);
        appendNewSaleToView(saleId, date, finalTotal, saleItems);
      }
    } catch (viewError) {
      // Silently fail if view update fails - don't break the sale submission
      Logger.log('Error updating Sales View: ' + viewError.toString());
    }

    // Build success message
    let message = `Sale recorded! Subtotal: $${subtotal.toFixed(2)}`;
    if (deliveryFee > 0) message += `, Delivery: -$${deliveryFee.toFixed(2)}`;
    if (saleDiscount > 0) message += `, Discount: -$${saleDiscount.toFixed(2)}`;
    message += `, Total: $${finalTotal.toFixed(2)}`;

    return {
      success: true,
      message: message,
    };
  } catch (error) {
    return { success: false, message: error.message };
  }
}

// =========================
// MONTHLY REPORT
// =========================

/**
 * Generate comprehensive monthly revenue report with insights
 */
function generateMonthlyReport() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let reportSheet = ss.getSheetByName('Report');
    if (!reportSheet) {
      reportSheet = ss.insertSheet('Report');
    }

    reportSheet.clear();

    const saleSheet = ss.getSheetByName('Sales');
    const itemSheet = ss.getSheetByName('SaleItems');

    if (!saleSheet || !itemSheet) {
      throw new Error('Sales or SaleItems sheet not found');
    }

    if (saleSheet.getLastRow() <= 1 || itemSheet.getLastRow() <= 1) {
      reportSheet.appendRow(['No sales data available']);
      SpreadsheetApp.getUi().alert('Report generated. No sales data found.');
      return;
    }

    // Get all sales with dates (get all columns to support new format)
    const salesData = saleSheet
      .getRange(2, 1, saleSheet.getLastRow() - 1, saleSheet.getLastColumn())
      .getValues();

    // Get sales headers to find column indices
    const saleHeaders = saleSheet
      .getRange(1, 1, 1, saleSheet.getLastColumn())
      .getValues()[0];
    const saleIdColIdx = 0;
    const saleDateColIdx = 1;
    const subtotalColIdx = saleHeaders.indexOf('Subtotal');
    const deliveryFeeColIdx = saleHeaders.indexOf('Delivery Fee');
    const discountColIdx = saleHeaders.indexOf('Discount');
    const totalAmountColIdx = saleHeaders.indexOf('Total Amount');

    // Get delivery costs data
    let deliveryCostsData = [];
    const deliverySheet = ss.getSheetByName('DeliveryCosts');
    if (deliverySheet && deliverySheet.getLastRow() > 1) {
      deliveryCostsData = deliverySheet
        .getRange(2, 1, deliverySheet.getLastRow() - 1, 3)
        .getValues();
    }

    // Get items data - adjust column count based on whether Product Name column exists
    const itemHeaders = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductName = itemHeaders.includes('Product Name');
    const itemColCount = hasProductName ? 7 : 6;
    const itemsData = itemSheet
      .getRange(2, 1, itemSheet.getLastRow() - 1, itemColCount)
      .getValues();

    // Column indices
    const productIdColIndex = 1;
    const productNameColIndex = hasProductName ? 2 : -1;
    const quantityColIndex = hasProductName ? 3 : 2;
    const isFreeColIndex = hasProductName ? 4 : 3;
    const unitPriceColIndex = hasProductName ? 5 : 4;
    const finalPriceColIndex = hasProductName ? 6 : 5;

    // Create map of saleId to date and amounts
    const saleDateMap = {};
    const saleAmountMap = {};
    const saleDeliveryMap = {};
    const saleDiscountMap = {};
    salesData.forEach((row) => {
      const saleId = row[saleIdColIdx];
      saleDateMap[saleId] = row[saleDateColIdx];
      // Use Total Amount column if exists, otherwise column 3 (old format)
      saleAmountMap[saleId] =
        totalAmountColIdx >= 0 ? row[totalAmountColIdx] : row[2];
      saleDeliveryMap[saleId] =
        deliveryFeeColIdx >= 0 ? row[deliveryFeeColIdx] || 0 : 0;
      saleDiscountMap[saleId] =
        discountColIdx >= 0 ? row[discountColIdx] || 0 : 0;
    });

    // Create map of delivery costs by month
    const monthlyDeliveryCosts = {};
    deliveryCostsData.forEach((row) => {
      const rawDate = row[1];
      if (rawDate) {
        const date = rawDate instanceof Date ? rawDate : new Date(rawDate);
        if (!isNaN(date.getTime())) {
          const year = date.getFullYear();
          const month = date.getMonth() + 1;
          const key = `${year}-${String(month).padStart(2, '0')}`;
          monthlyDeliveryCosts[key] =
            (monthlyDeliveryCosts[key] || 0) + (row[2] || 0);
        }
      }
    });

    // Debug logging
    Logger.log('Sales data count: ' + salesData.length);
    Logger.log('Items data count: ' + itemsData.length);
    Logger.log(
      'Sample sale IDs from Sales: ' +
        Object.keys(saleDateMap).slice(0, 3).join(', '),
    );
    Logger.log(
      'Sample sale IDs from Items: ' +
        itemsData
          .slice(0, 3)
          .map((r) => r[0])
          .join(', '),
    );
    if (salesData.length > 0) {
      Logger.log('First sale row: ' + JSON.stringify(salesData[0]));
    }
    if (itemsData.length > 0) {
      Logger.log('First item row: ' + JSON.stringify(itemsData[0]));
    }

    // Aggregate data by month
    const monthDataMap = {};

    // Aggregate data by product
    const productDataMap = {};

    // Track day of week sales
    const dayOfWeekSales = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 }; // Sun-Sat
    const dayOfWeekCount = { 0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0 };

    // Track free items given
    let totalFreeItems = 0;
    let totalFreeItemsValue = 0;

    let matchedItems = 0;
    let unmatchedItems = 0;

    // Process items for revenue and quantity
    itemsData.forEach((row) => {
      const saleId = row[0];
      const productId = row[productIdColIndex];
      const productName = hasProductName ? row[productNameColIndex] : productId;
      const quantity = row[quantityColIndex] || 0;
      const isFree = row[isFreeColIndex];
      const unitPrice = row[unitPriceColIndex] || 0;
      const finalPrice = row[finalPriceColIndex] || 0;
      const rawDate = saleDateMap[saleId];

      if (!rawDate) {
        unmatchedItems++;
        Logger.log('No date found for saleId: ' + saleId);
        return;
      }

      // Ensure date is a Date object
      let date;
      if (rawDate instanceof Date) {
        date = rawDate;
      } else {
        // Try parsing string date in format "DD/MM/YYYY, HH:MM:SS"
        const dateStr = String(rawDate);
        const match = dateStr.match(
          /(\d{1,2})\/(\d{1,2})\/(\d{4}),?\s*(\d{1,2}):(\d{2}):(\d{2})/,
        );
        if (match) {
          const [, day, month, year, hour, min, sec] = match;
          date = new Date(year, month - 1, day, hour, min, sec);
        } else {
          // Fallback to standard parsing
          date = new Date(rawDate);
        }
      }

      if (isNaN(date.getTime())) {
        Logger.log(
          'Invalid date for saleId: ' + saleId + ', rawDate: ' + rawDate,
        );
        unmatchedItems++;
        return;
      }

      matchedItems++;
      const year = date.getFullYear();
      const month = date.getMonth() + 1;
      const key = `${year}-${String(month).padStart(2, '0')}`;
      const dayOfWeek = date.getDay();

      // Monthly aggregation
      if (!monthDataMap[key]) {
        monthDataMap[key] = {
          revenue: 0,
          itemsSold: 0,
          transactions: new Set(),
          freeItems: 0,
          deliveryFees: 0,
          discounts: 0,
          processedSales: new Set(),
        };
      }
      monthDataMap[key].revenue += finalPrice;
      monthDataMap[key].itemsSold += quantity;
      monthDataMap[key].transactions.add(saleId);
      if (isFree) {
        monthDataMap[key].freeItems += quantity;
      }
      // Add delivery and discount only once per sale
      if (!monthDataMap[key].processedSales.has(saleId)) {
        monthDataMap[key].deliveryFees += saleDeliveryMap[saleId] || 0;
        monthDataMap[key].discounts += saleDiscountMap[saleId] || 0;
        monthDataMap[key].processedSales.add(saleId);
      }

      // Product aggregation
      const productKey = productName || productId;
      if (!productDataMap[productKey]) {
        productDataMap[productKey] = {
          productId: productId,
          productName: productName,
          totalQuantity: 0,
          totalRevenue: 0,
          freeQuantity: 0,
          transactionCount: new Set(),
          avgUnitPrice: 0,
          priceSum: 0,
          paidQuantity: 0,
        };
      }
      productDataMap[productKey].totalQuantity += quantity;
      productDataMap[productKey].totalRevenue += finalPrice;
      productDataMap[productKey].transactionCount.add(saleId);
      if (isFree) {
        productDataMap[productKey].freeQuantity += quantity;
        totalFreeItems += quantity;
      } else {
        productDataMap[productKey].priceSum += unitPrice * quantity;
        productDataMap[productKey].paidQuantity += quantity;
      }

      // Day of week tracking
      dayOfWeekSales[dayOfWeek] += finalPrice;
      dayOfWeekCount[dayOfWeek]++;
    });

    // Log summary
    Logger.log(
      'Matched items: ' + matchedItems + ', Unmatched items: ' + unmatchedItems,
    );
    Logger.log('Month data keys: ' + Object.keys(monthDataMap).join(', '));
    Logger.log('Product data keys: ' + Object.keys(productDataMap).join(', '));

    // Calculate average unit price for each product
    Object.keys(productDataMap).forEach((key) => {
      const prod = productDataMap[key];
      prod.avgUnitPrice =
        prod.paidQuantity > 0 ? prod.priceSum / prod.paidQuantity : 0;
    });

    // Sort by year-month
    const sortedKeys = Object.keys(monthDataMap).sort();

    let currentRow = 1;

    // ==================== SECTION 1: MONTHLY OVERVIEW ====================
    reportSheet.getRange(currentRow, 1).setValue('📊 MONTHLY SALES OVERVIEW');
    reportSheet.getRange(currentRow, 1, 1, 10).merge();
    reportSheet
      .getRange(currentRow, 1)
      .setFontWeight('bold')
      .setFontSize(14)
      .setBackground('#1a73e8')
      .setFontColor('#ffffff');
    currentRow += 2;

    // Write header
    const monthlyHeaders = [
      'Year',
      'Month',
      'Gross Revenue',
      'Delivery Fees',
      'Discounts',
      'Net Revenue',
      'Transactions',
      'Items Sold',
      'Avg Sale',
      'Growth %',
    ];
    reportSheet.getRange(currentRow, 1, 1, 10).setValues([monthlyHeaders]);
    reportSheet
      .getRange(currentRow, 1, 1, 10)
      .setFontWeight('bold')
      .setBackground('#4285f4')
      .setFontColor('#ffffff');
    currentRow++;

    // Write monthly data
    let previousNetRevenue = null;
    sortedKeys.forEach((key) => {
      const [year, month] = key.split('-');
      const data = monthDataMap[key];
      const grossRevenue = data.revenue;
      const deliveryFees = data.deliveryFees || 0;
      const discounts = data.discounts || 0;
      const netRevenue = grossRevenue - deliveryFees - discounts;
      const transactions = data.transactions.size;
      const itemsSold = data.itemsSold;
      const avgSale = transactions > 0 ? netRevenue / transactions : 0;

      // Calculate growth percentage based on net revenue
      let growth = '';
      if (previousNetRevenue !== null && previousNetRevenue > 0) {
        const growthPercent =
          ((netRevenue - previousNetRevenue) / previousNetRevenue) * 100;
        growth =
          growthPercent >= 0
            ? `+${growthPercent.toFixed(1)}%`
            : `${growthPercent.toFixed(1)}%`;
      } else {
        growth = 'N/A';
      }

      reportSheet
        .getRange(currentRow, 1, 1, 10)
        .setValues([
          [
            year,
            month,
            grossRevenue.toFixed(2),
            deliveryFees.toFixed(2),
            discounts.toFixed(2),
            netRevenue.toFixed(2),
            transactions,
            itemsSold,
            avgSale.toFixed(2),
            growth,
          ],
        ]);
      currentRow++;

      previousNetRevenue = netRevenue;
    });

    // Add borders to monthly section
    const monthlyDataRange = reportSheet.getRange(
      currentRow - sortedKeys.length - 1,
      1,
      sortedKeys.length + 1,
      10,
    );
    monthlyDataRange.setBorder(true, true, true, true, true, true);

    currentRow += 2;

    // ==================== SECTION 2: PRODUCT PERFORMANCE ====================
    reportSheet
      .getRange(currentRow, 1)
      .setValue('📦 PRODUCT PERFORMANCE (Top Sellers)');
    reportSheet.getRange(currentRow, 1, 1, 8).merge();
    reportSheet
      .getRange(currentRow, 1)
      .setFontWeight('bold')
      .setFontSize(14)
      .setBackground('#34a853')
      .setFontColor('#ffffff');
    currentRow += 2;

    // Sort products by quantity sold (descending)
    const sortedProducts = Object.values(productDataMap).sort(
      (a, b) => b.totalQuantity - a.totalQuantity,
    );

    // Product headers
    const productHeaders = [
      'Rank',
      'Product Name',
      'Units Sold',
      'Revenue',
      'Free Given',
      'Avg Price',
      'Times Ordered',
      '% of Sales',
    ];
    reportSheet.getRange(currentRow, 1, 1, 8).setValues([productHeaders]);
    reportSheet
      .getRange(currentRow, 1, 1, 8)
      .setFontWeight('bold')
      .setBackground('#34a853')
      .setFontColor('#ffffff');
    currentRow++;

    const totalUnits = sortedProducts.reduce(
      (sum, p) => sum + p.totalQuantity,
      0,
    );
    const productStartRow = currentRow;

    sortedProducts.forEach((prod, index) => {
      const percentOfSales =
        totalUnits > 0
          ? ((prod.totalQuantity / totalUnits) * 100).toFixed(1) + '%'
          : '0%';
      reportSheet
        .getRange(currentRow, 1, 1, 8)
        .setValues([
          [
            index + 1,
            prod.productName || prod.productId,
            prod.totalQuantity,
            prod.totalRevenue.toFixed(2),
            prod.freeQuantity,
            prod.avgUnitPrice.toFixed(2),
            prod.transactionCount.size,
            percentOfSales,
          ],
        ]);
      currentRow++;
    });

    // Add borders to product section
    const productDataRange = reportSheet.getRange(
      productStartRow - 1,
      1,
      sortedProducts.length + 1,
      8,
    );
    productDataRange.setBorder(true, true, true, true, true, true);

    currentRow += 2;

    // ==================== SECTION 3: SALES BY DAY OF WEEK ====================
    reportSheet.getRange(currentRow, 1).setValue('📅 SALES BY DAY OF WEEK');
    reportSheet.getRange(currentRow, 1, 1, 4).merge();
    reportSheet
      .getRange(currentRow, 1)
      .setFontWeight('bold')
      .setFontSize(14)
      .setBackground('#fbbc04')
      .setFontColor('#000000');
    currentRow += 2;

    const dayNames = [
      'Sunday',
      'Monday',
      'Tuesday',
      'Wednesday',
      'Thursday',
      'Friday',
      'Saturday',
    ];
    reportSheet
      .getRange(currentRow, 1, 1, 4)
      .setValues([['Day', 'Revenue', 'Orders', 'Avg Order']]);
    reportSheet
      .getRange(currentRow, 1, 1, 4)
      .setFontWeight('bold')
      .setBackground('#fbbc04');
    currentRow++;

    const dayStartRow = currentRow;
    let bestDay = { day: '', revenue: 0 };

    dayNames.forEach((dayName, index) => {
      const revenue = dayOfWeekSales[index];
      const count = dayOfWeekCount[index];
      const avgOrder = count > 0 ? revenue / count : 0;

      if (revenue > bestDay.revenue) {
        bestDay = { day: dayName, revenue: revenue };
      }

      reportSheet
        .getRange(currentRow, 1, 1, 4)
        .setValues([[dayName, revenue.toFixed(2), count, avgOrder.toFixed(2)]]);
      currentRow++;
    });

    // Add borders to day section
    const dayDataRange = reportSheet.getRange(dayStartRow - 1, 1, 8, 4);
    dayDataRange.setBorder(true, true, true, true, true, true);

    currentRow += 2;

    // ==================== SECTION 4: KEY INSIGHTS & SUMMARY ====================
    reportSheet.getRange(currentRow, 1).setValue('💡 KEY INSIGHTS & SUMMARY');
    reportSheet.getRange(currentRow, 1, 1, 4).merge();
    reportSheet
      .getRange(currentRow, 1)
      .setFontWeight('bold')
      .setFontSize(14)
      .setBackground('#ea4335')
      .setFontColor('#ffffff');
    currentRow += 2;

    // Calculate totals and insights
    let totalGrossRevenue = 0;
    let totalDeliveryFees = 0;
    let totalDiscounts = 0;
    let totalTransactions = 0;
    let totalItemsSold = 0;
    let totalFreeGiven = 0;

    sortedKeys.forEach((key) => {
      const data = monthDataMap[key];
      totalGrossRevenue += data.revenue;
      totalDeliveryFees += data.deliveryFees || 0;
      totalDiscounts += data.discounts || 0;
      totalTransactions += data.transactions.size;
      totalItemsSold += data.itemsSold;
      totalFreeGiven += data.freeItems || 0;
    });

    const totalNetRevenue =
      totalGrossRevenue - totalDeliveryFees - totalDiscounts;
    const overallAvgSale =
      totalTransactions > 0 ? totalNetRevenue / totalTransactions : 0;
    const avgItemsPerTransaction =
      totalTransactions > 0 ? totalItemsSold / totalTransactions : 0;

    // Get top 3 products
    const top3Products = sortedProducts
      .slice(0, 3)
      .map((p) => p.productName || p.productId)
      .join(', ');

    // Get slowest selling product (exclude products with 0 sales)
    const slowestProduct =
      sortedProducts.length > 0
        ? sortedProducts[sortedProducts.length - 1]
        : null;

    // Calculate revenue per item
    const revenuePerItem =
      totalItemsSold > 0 ? totalNetRevenue / totalItemsSold : 0;

    const insights = [
      ['📈 Gross Revenue:', '$' + totalGrossRevenue.toFixed(2)],
      ['🚚 Total Delivery Fees:', '-$' + totalDeliveryFees.toFixed(2)],
      ['🏷️ Total Discounts:', '-$' + totalDiscounts.toFixed(2)],
      ['💵 Net Revenue:', '$' + totalNetRevenue.toFixed(2)],
      [''],
      ['🛒 Total Transactions:', totalTransactions],
      ['📦 Total Items Sold:', totalItemsSold],
      ['🎁 Free Items Given:', totalFreeGiven],
      ['💰 Average Sale Value:', '$' + overallAvgSale.toFixed(2)],
      ['📊 Avg Items per Transaction:', avgItemsPerTransaction.toFixed(1)],
      ['💵 Revenue per Item:', '$' + revenuePerItem.toFixed(2)],
      ['📅 Reporting Period:', sortedKeys.length + ' months'],
      [''],
      ['🏆 TOP PERFORMERS:'],
      ['Best Selling Products:', top3Products || 'N/A'],
      [
        'Best Sales Day:',
        bestDay.day + ' ($' + bestDay.revenue.toFixed(2) + ')',
      ],
      [''],
      ['⚠️ ATTENTION NEEDED:'],
      [
        'Slowest Selling:',
        slowestProduct
          ? (slowestProduct.productName || slowestProduct.productId) +
            ' (' +
            slowestProduct.totalQuantity +
            ' units)'
          : 'N/A',
      ],
      ['Free Items Cost:', totalFreeGiven + ' items given away'],
    ];

    const insightsStartRow = currentRow;
    insights.forEach((row) => {
      if (row.length === 2) {
        reportSheet
          .getRange(currentRow, 1)
          .setValue(row[0])
          .setFontWeight('bold');
        const valueCell = reportSheet.getRange(currentRow, 2);
        valueCell.setValue(row[1]);

        // Set number format for non-currency values
        const label = row[0];
        if (
          label.includes('Total Items Sold') ||
          label.includes('Free Items Given') ||
          label.includes('Total Transactions') ||
          label.includes('Avg Items per Transaction')
        ) {
          valueCell.setNumberFormat('0');
        } else if (label.includes('Reporting Period')) {
          valueCell.setNumberFormat('@'); // Text format
        }
      } else if (
        (row[0] && row[0].includes('TOP')) ||
        (row[0] && row[0].includes('ATTENTION'))
      ) {
        reportSheet
          .getRange(currentRow, 1)
          .setValue(row[0])
          .setFontWeight('bold')
          .setFontSize(11);
      }
      currentRow++;
    });

    // Set column widths
    reportSheet.setColumnWidth(1, 80);
    reportSheet.setColumnWidth(2, 180);
    reportSheet.setColumnWidth(3, 100);
    reportSheet.setColumnWidth(4, 100);
    reportSheet.setColumnWidth(5, 100);
    reportSheet.setColumnWidth(6, 100);
    reportSheet.setColumnWidth(7, 110);
    reportSheet.setColumnWidth(8, 100);

    SpreadsheetApp.getUi().alert(
      "Enhanced report generated successfully in the 'Report' sheet!\n\nIncludes:\n• Monthly Overview\n• Product Performance\n• Sales by Day of Week\n• Key Insights & Summary",
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error generating report: ' + error.message);
    Logger.log('Error details: ' + error.toString());
  }
}

/**
 * Helper function to get product name by ID
 */
function getProductNameById(productId) {
  const products = getProducts();
  const product = products.find((p) => p.id === productId);
  return product ? product.name : 'Unknown Product';
}

/**
 * Update existing SaleItems to include Product Names
 * Run this once to populate product names for existing records
 */
function updateSaleItemsWithProductNames() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const itemSheet = ss.getSheetByName('SaleItems');

    if (!itemSheet || itemSheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert('No SaleItems data found.');
      return;
    }

    // Get headers to find column indices
    const headers = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const productIdColIndex = headers.indexOf('Product ID') + 1; // Column B (index 1, so +1 = 2)
    const productNameColIndex = headers.indexOf('Product Name') + 1;

    if (productIdColIndex === 0) {
      SpreadsheetApp.getUi().alert(
        'Error: Product ID column not found in SaleItems sheet.',
      );
      return;
    }

    if (productNameColIndex === 0) {
      // Insert Product Name column after Product ID
      itemSheet.insertColumnAfter(productIdColIndex);
      itemSheet.getRange(1, productIdColIndex + 1).setValue('Product Name');
      // Update column index
      const newProductNameColIndex = productIdColIndex + 1;

      // Populate product names for all rows
      const rows = itemSheet
        .getRange(2, 1, itemSheet.getLastRow() - 1, itemSheet.getLastColumn())
        .getValues();
      rows.forEach((row, index) => {
        const productId = row[productIdColIndex - 1]; // Product ID column (0-indexed)
        const productName = getProductNameById(productId);
        itemSheet
          .getRange(index + 2, newProductNameColIndex)
          .setValue(productName);
      });
    } else {
      // Column exists, update empty cells or all cells
      const rows = itemSheet
        .getRange(2, 1, itemSheet.getLastRow() - 1, itemSheet.getLastColumn())
        .getValues();
      let updatedCount = 0;

      rows.forEach((row, index) => {
        const productId = row[productIdColIndex - 1]; // Product ID column (0-indexed)
        const currentProductName = row[productNameColIndex - 1];

        // Update if product name is empty or update all to ensure accuracy
        if (
          !currentProductName ||
          currentProductName === '' ||
          currentProductName === 'Unknown Product'
        ) {
          const productName = getProductNameById(productId);
          itemSheet
            .getRange(index + 2, productNameColIndex)
            .setValue(productName);
          updatedCount++;
        }
      });

      if (updatedCount === 0) {
        SpreadsheetApp.getUi().alert(
          'All SaleItems already have product names.',
        );
      } else {
        SpreadsheetApp.getUi().alert(
          `Updated ${updatedCount} SaleItems with product names.`,
        );
      }
      return;
    }

    SpreadsheetApp.getUi().alert(
      'SaleItems updated with product names successfully!',
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error updating SaleItems: ' + error.message);
    Logger.log('Error details: ' + error.toString());
  }
}

// =========================
// VIEW SALES FUNCTIONS
// =========================

/**
 * Get all sales with their items combined
 */
function getAllSalesWithItems() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const saleSheet = ss.getSheetByName('Sales');
    const itemSheet = ss.getSheetByName('SaleItems');

    if (!saleSheet || !itemSheet) {
      return [];
    }

    if (saleSheet.getLastRow() <= 1 || itemSheet.getLastRow() <= 1) {
      return [];
    }

    // Get all sales
    const salesData = saleSheet
      .getRange(2, 1, saleSheet.getLastRow() - 1, 3)
      .getValues();

    // Get all items
    const itemHeaders = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductName = itemHeaders.includes('Product Name');
    const itemColCount = hasProductName ? 7 : 6;
    const itemsData = itemSheet
      .getRange(2, 1, itemSheet.getLastRow() - 1, itemColCount)
      .getValues();

    // Column indices
    const productNameColIndex = hasProductName ? 2 : -1;
    const quantityColIndex = hasProductName ? 3 : 2;
    const isFreeColIndex = hasProductName ? 4 : 3;
    const unitPriceColIndex = hasProductName ? 5 : 4;
    const finalPriceColIndex = hasProductName ? 6 : 5;

    // Group items by sale ID
    const itemsBySale = {};
    itemsData.forEach((row) => {
      const saleId = row[0];
      if (!itemsBySale[saleId]) {
        itemsBySale[saleId] = [];
      }
      itemsBySale[saleId].push({
        productId: row[1],
        productName: hasProductName ? row[productNameColIndex] : 'N/A',
        quantity: row[quantityColIndex],
        isFree: row[isFreeColIndex],
        unitPrice: row[unitPriceColIndex],
        finalPrice: row[finalPriceColIndex],
      });
    });

    // Combine sales with their items
    const salesWithItems = salesData.map((row) => {
      const saleId = row[0];
      const date = row[1];
      const totalAmount = row[2];
      // Convert date to ISO string for proper JSON serialization
      const dateString =
        date instanceof Date
          ? date.toISOString()
          : date
          ? new Date(date).toISOString()
          : new Date().toISOString();
      return {
        saleId: saleId,
        date: dateString, // Store as ISO string for JSON compatibility
        totalAmount: totalAmount,
        items: itemsBySale[saleId] || [],
      };
    });

    // Sort by date (newest first)
    salesWithItems.sort((a, b) => {
      return new Date(b.date) - new Date(a.date);
    });

    return salesWithItems;
  } catch (error) {
    Logger.log('Error getting sales with items: ' + error.toString());
    return [];
  }
}

/**
 * Get sale details by sale ID
 */
function getSaleDetails(saleId) {
  try {
    const allSales = getAllSalesWithItems();
    return allSales.find((sale) => sale.saleId === saleId) || null;
  } catch (error) {
    Logger.log('Error getting sale details: ' + error.toString());
    return null;
  }
}

/**
 * Open sidebar to view sale details
 */
function openViewSaleSidebar() {
  const template = HtmlService.createTemplateFromFile('viewSale');
  const salesData = getAllSalesWithItems();
  Logger.log('Sales data retrieved:', salesData.length, 'sales');
  template.sales = salesData;
  const html = template.evaluate().setTitle('View Sale Details').setWidth(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Append new sale to Sales View sheet (optimized - only adds new data, doesn't regenerate)
 */
function appendNewSaleToView(saleId, saleDate, totalAmount, saleItems) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const viewSheet = ss.getSheetByName('Sales View');
    const itemSheet = ss.getSheetByName('SaleItems');

    if (!viewSheet || !itemSheet) {
      return;
    }

    // Check if sheet has header (if not, it's empty, so populate it)
    if (viewSheet.getLastRow() === 0) {
      // Sheet is empty, populate it fully
      populateSalesViewSheet(viewSheet);
      return;
    }

    // Get product names for the sale items
    const products = getProducts();
    const productMap = {};
    products.forEach((p) => {
      productMap[p.id] = p.name;
    });

    // Format the sale date
    const saleDateObj =
      saleDate instanceof Date ? saleDate : new Date(saleDate);
    const formattedDate = saleDateObj.toLocaleString();

    // Get the items for this sale from SaleItems sheet (just added)
    const itemHeaders = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductName = itemHeaders.includes('Product Name');
    const itemColCount = hasProductName ? 7 : 6;

    // Only get the last few rows (the new sale items we just added)
    const lastRow = itemSheet.getLastRow();
    const recentRows = Math.min(10, lastRow - 1); // Check last 10 rows or all if less
    const itemsData = itemSheet
      .getRange(
        Math.max(2, lastRow - recentRows + 1),
        1,
        recentRows,
        itemColCount,
      )
      .getValues();

    const productNameColIndex = hasProductName ? 2 : -1;
    const quantityColIndex = hasProductName ? 3 : 2;
    const isFreeColIndex = hasProductName ? 4 : 3;
    const unitPriceColIndex = hasProductName ? 5 : 4;
    const finalPriceColIndex = hasProductName ? 6 : 5;

    // Find items for this sale
    const saleItemsData = itemsData.filter((row) => row[0] === saleId);

    if (saleItemsData.length === 0) {
      // No items found, skip (might be a timing issue, fallback to full update)
      return;
    }

    // Insert new sale at row 2 (after header, newest first)
    let insertRow = 2;

    if (saleItemsData.length === 0) {
      // Sale with no items
      viewSheet.insertRowBefore(insertRow);
      viewSheet.getRange(insertRow, 1).setValue(saleId);
      viewSheet.getRange(insertRow, 2).setValue(formattedDate);
      viewSheet.getRange(insertRow, 3).setValue('No items');
      viewSheet.getRange(insertRow, 7).setValue(totalAmount.toFixed(2));
      viewSheet.getRange(insertRow, 7).setFontWeight('bold');
      viewSheet.getRange(insertRow, 7).setBackground('#e8f0fe');
    } else {
      // Add each item
      saleItemsData.forEach((itemRow, index) => {
        const isFirstItem = index === 0;
        viewSheet.insertRowBefore(insertRow);

        if (isFirstItem) {
          viewSheet.getRange(insertRow, 1).setValue(saleId);
          viewSheet.getRange(insertRow, 2).setValue(formattedDate);
          viewSheet.getRange(insertRow, 7).setValue(totalAmount.toFixed(2));
          viewSheet.getRange(insertRow, 7).setFontWeight('bold');
          viewSheet.getRange(insertRow, 7).setBackground('#e8f0fe');
        }

        const productName = hasProductName
          ? itemRow[productNameColIndex]
          : productMap[itemRow[1]] || 'Unknown';
        const quantity = itemRow[quantityColIndex];
        const isFree = itemRow[isFreeColIndex];
        const unitPrice = itemRow[unitPriceColIndex];
        const finalPrice = itemRow[finalPriceColIndex];

        viewSheet.getRange(insertRow, 3).setValue(productName);
        viewSheet.getRange(insertRow, 4).setValue(quantity);
        viewSheet
          .getRange(insertRow, 5)
          .setValue(isFree ? 'FREE' : unitPrice.toFixed(2));
        viewSheet.getRange(insertRow, 6).setValue(finalPrice.toFixed(2));

        // Format number columns
        if (!isFree) {
          viewSheet.getRange(insertRow, 5).setNumberFormat('#,##0.00');
          viewSheet.getRange(insertRow, 6).setNumberFormat('#,##0.00');
        }

        insertRow++;
      });
    }

    // Add borders to new rows
    const numRows = saleItemsData.length > 0 ? saleItemsData.length : 1;
    const newRange = viewSheet.getRange(insertRow - numRows, 1, numRows, 7);
    newRange.setBorder(true, true, true, true, true, true);
  } catch (error) {
    Logger.log('Error appending to Sales View: ' + error.toString());
    // Silently fail - don't break sale submission
  }
}

/**
 * Update the Sales View sheet (internal function, called automatically)
 * Only used as fallback or when sheet is empty
 */
function updateSalesViewSheet() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const viewSheet = ss.getSheetByName('Sales View');

    if (!viewSheet) {
      return; // Sheet doesn't exist yet, skip update
    }

    viewSheet.clear();
    populateSalesViewSheet(viewSheet);
  } catch (error) {
    Logger.log('Error updating Sales View: ' + error.toString());
  }
}

/**
 * Populate the Sales View sheet with data (shared logic)
 */
function populateSalesViewSheet(viewSheet) {
  const salesWithItems = getAllSalesWithItems();

  if (salesWithItems.length === 0) {
    viewSheet.appendRow(['No sales data available']);
    return;
  }

  // Create header
  viewSheet.appendRow([
    'Sale ID',
    'Date',
    'Product Name',
    'Quantity',
    'Unit Price',
    'Item Total',
    'Sale Total',
  ]);

  // Add data
  salesWithItems.forEach((sale) => {
    const saleDateObj =
      sale.date instanceof Date ? sale.date : new Date(sale.date);
    const saleDate = saleDateObj.toLocaleString(); // Includes both date and time

    if (sale.items.length === 0) {
      // Sale with no items
      viewSheet.appendRow([
        sale.saleId,
        saleDate,
        'No items',
        '',
        '',
        '',
        sale.totalAmount.toFixed(2),
      ]);
    } else {
      // Add each item
      sale.items.forEach((item, index) => {
        const isFirstItem = index === 0;
        viewSheet.appendRow([
          isFirstItem ? sale.saleId : '', // Only show sale ID on first row
          isFirstItem ? saleDate : '', // Only show date on first row
          item.productName,
          item.quantity,
          item.isFree ? 'FREE' : item.unitPrice.toFixed(2),
          item.finalPrice.toFixed(2),
          isFirstItem ? sale.totalAmount.toFixed(2) : '', // Only show total on first row
        ]);
      });
    }
  });

  // Format the sheet
  const headerRange = viewSheet.getRange(1, 1, 1, 7);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#4285f4');
  headerRange.setFontColor('#ffffff');

  // Set column widths
  viewSheet.setColumnWidth(1, 120); // Sale ID
  viewSheet.setColumnWidth(2, 100); // Date
  viewSheet.setColumnWidth(3, 200); // Product Name
  viewSheet.setColumnWidth(4, 80); // Quantity
  viewSheet.setColumnWidth(5, 100); // Unit Price
  viewSheet.setColumnWidth(6, 100); // Item Total
  viewSheet.setColumnWidth(7, 100); // Sale Total

  // Add borders
  const dataRange = viewSheet.getRange(1, 1, viewSheet.getLastRow(), 7);
  dataRange.setBorder(true, true, true, true, true, true);

  // Format numbers
  viewSheet
    .getRange(2, 5, viewSheet.getLastRow() - 1, 3)
    .setNumberFormat('#,##0.00');

  // Highlight sale total rows
  let currentRow = 2;
  salesWithItems.forEach((sale) => {
    if (sale.items.length > 0) {
      const firstItemRow = currentRow;
      viewSheet
        .getRange(firstItemRow, 7, 1, 1)
        .setFontWeight('bold')
        .setBackground('#e8f0fe');
      currentRow += sale.items.length;
    } else {
      viewSheet
        .getRange(currentRow, 7, 1, 1)
        .setFontWeight('bold')
        .setBackground('#e8f0fe');
      currentRow += 1;
    }
  });
}

/**
 * Create a combined sales view sheet with all sales and their items
 */
function createCombinedSalesView() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let viewSheet = ss.getSheetByName('Sales View');

    if (!viewSheet) {
      viewSheet = ss.insertSheet('Sales View');
    } else {
      viewSheet.clear();
    }

    populateSalesViewSheet(viewSheet);

    SpreadsheetApp.getUi().alert(
      "Combined Sales View created/updated successfully in the 'Sales View' sheet!",
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error creating sales view: ' + error.message);
    Logger.log('Error details: ' + error.toString());
  }
}

/**
 * Sync Sales View sheet back to Sales and SaleItems sheets
 * This reads from Sales View and updates the underlying data sheets
 */
function syncSalesViewToData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const viewSheet = ss.getSheetByName('Sales View');
    const saleSheet = ss.getSheetByName('Sales');
    const itemSheet = ss.getSheetByName('SaleItems');

    if (!viewSheet) {
      SpreadsheetApp.getUi().alert(
        'Sales View sheet not found. Please create it first using "Create Combined Sales View".',
      );
      return;
    }

    if (!saleSheet || !itemSheet) {
      SpreadsheetApp.getUi().alert('Sales or SaleItems sheet not found.');
      return;
    }

    if (viewSheet.getLastRow() <= 1) {
      SpreadsheetApp.getUi().alert('Sales View sheet is empty.');
      return;
    }

    // Read Sales View data (skip header)
    // Columns: Sale ID, Date, Product Name, Quantity, Unit Price, Item Total, Sale Total
    const viewData = viewSheet
      .getRange(2, 1, viewSheet.getLastRow() - 1, 7)
      .getValues();

    // Get product map for looking up product IDs by name
    const products = getProducts();
    const productNameToId = {};
    products.forEach((p) => {
      productNameToId[p.name] = p.id;
    });

    // Parse Sales View into sales and items
    const salesMap = {}; // saleId -> { date, totalAmount, subtotal }
    const itemsMap = {}; // saleId -> [ { productName, quantity, unitPrice, itemTotal } ]

    let currentSaleId = null;
    let currentDate = null;

    viewData.forEach((row) => {
      const saleId = row[0];
      const date = row[1];
      const productName = row[2];
      const quantity = row[3];
      const unitPrice = row[4];
      const itemTotal = row[5];
      const saleTotal = row[6];

      // If sale ID is present, this is a new sale or first item of a sale
      if (saleId && saleId !== '') {
        currentSaleId = saleId;
        currentDate = date;

        if (!salesMap[currentSaleId]) {
          const total =
            typeof saleTotal === 'number'
              ? saleTotal
              : parseFloat(saleTotal) || 0;
          salesMap[currentSaleId] = {
            date: currentDate,
            totalAmount: total,
            subtotal: 0, // Will be calculated from items
          };
          itemsMap[currentSaleId] = [];
        } else if (saleTotal && saleTotal !== '') {
          // Update total if provided
          const total =
            typeof saleTotal === 'number'
              ? saleTotal
              : parseFloat(saleTotal) || 0;
          salesMap[currentSaleId].totalAmount = total;
        }
      }

      // Add item if product name exists and it's not "No items"
      if (
        currentSaleId &&
        productName &&
        productName !== '' &&
        productName !== 'No items'
      ) {
        const isFree =
          unitPrice === 'FREE' || unitPrice === 0 || unitPrice === '0';
        const actualUnitPrice = isFree
          ? 0
          : typeof unitPrice === 'number'
          ? unitPrice
          : parseFloat(unitPrice) || 0;
        const actualItemTotal =
          typeof itemTotal === 'number'
            ? itemTotal
            : parseFloat(itemTotal) || 0;

        itemsMap[currentSaleId].push({
          productName: productName,
          productId: productNameToId[productName] || '',
          quantity: quantity || 0,
          isFree: isFree,
          unitPrice: actualUnitPrice,
          finalPrice: actualItemTotal,
        });
        // Add to subtotal
        salesMap[currentSaleId].subtotal += actualItemTotal;
      }
    });

    // Calculate delivery fees and discounts for each sale
    Object.keys(salesMap).forEach((saleId) => {
      const sale = salesMap[saleId];
      const difference = sale.subtotal - sale.totalAmount;
      // If total is less than subtotal, the difference could be delivery fee or discount
      // For now, we'll assume it's delivery fee if positive
      sale.deliveryFee = difference > 0 ? difference : 0;
      sale.discount = 0; // Discount would need to be manually entered or calculated separately
    });

    // Check if Sales sheet has new format columns
    const saleHeaders = saleSheet
      .getRange(1, 1, 1, saleSheet.getLastColumn())
      .getValues()[0];
    const hasNewFormat = saleHeaders.includes('Subtotal');

    // Clean up orphaned rows (rows without Sale ID)
    if (saleSheet.getLastRow() > 1) {
      const allSalesData = saleSheet
        .getRange(2, 1, saleSheet.getLastRow() - 1, saleSheet.getLastColumn())
        .getValues();
      const rowsToDelete = [];
      for (let i = 0; i < allSalesData.length; i++) {
        if (!allSalesData[i][0] || allSalesData[i][0] === '') {
          rowsToDelete.push(i + 2); // +2 because row 1 is header
        }
      }
      // Delete from bottom to top
      rowsToDelete.sort((a, b) => b - a);
      rowsToDelete.forEach((rowNum) => {
        saleSheet.deleteRow(rowNum);
      });
    }

    // Clear and rebuild Sales sheet (keep header)
    if (saleSheet.getLastRow() > 1) {
      const colsToClear = hasNewFormat ? saleSheet.getLastColumn() : 3;
      saleSheet.getRange(2, 1, saleSheet.getLastRow() - 1, colsToClear).clear();
    }

    // Clear and rebuild SaleItems sheet (keep header)
    const itemHeaders = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductNameCol = itemHeaders.includes('Product Name');
    const itemColCount = hasProductNameCol ? 7 : 6;

    if (itemSheet.getLastRow() > 1) {
      itemSheet
        .getRange(2, 1, itemSheet.getLastRow() - 1, itemColCount)
        .clear();
    }

    // Write sales data
    const saleIds = Object.keys(salesMap);
    let saleRow = 2;
    let itemRow = 2;

    saleIds.forEach((saleId) => {
      const sale = salesMap[saleId];
      const items = itemsMap[saleId] || [];

      // Write to Sales sheet
      saleSheet.getRange(saleRow, 1).setValue(saleId);
      saleSheet.getRange(saleRow, 2).setValue(sale.date);

      if (hasNewFormat) {
        // New format: Sale ID, Date, Subtotal, Delivery Fee, Discount, Total Amount
        const subtotalCol = saleHeaders.indexOf('Subtotal') + 1;
        const deliveryFeeCol = saleHeaders.indexOf('Delivery Fee') + 1;
        const discountCol = saleHeaders.indexOf('Discount') + 1;
        const totalCol = saleHeaders.indexOf('Total Amount') + 1;

        if (subtotalCol > 0)
          saleSheet.getRange(saleRow, subtotalCol).setValue(sale.subtotal);
        if (deliveryFeeCol > 0)
          saleSheet
            .getRange(saleRow, deliveryFeeCol)
            .setValue(sale.deliveryFee);
        if (discountCol > 0)
          saleSheet.getRange(saleRow, discountCol).setValue(sale.discount);
        if (totalCol > 0) {
          saleSheet.getRange(saleRow, totalCol).setValue(sale.totalAmount);
        }
      } else {
        // Old format: Sale ID, Date, Total Amount
        saleSheet.getRange(saleRow, 3).setValue(sale.totalAmount);
      }
      saleRow++;

      // Update DeliveryCosts sheet if there's a delivery fee
      if (sale.deliveryFee > 0) {
        let deliverySheet = ss.getSheetByName('DeliveryCosts');
        if (!deliverySheet) {
          deliverySheet = ss.insertSheet('DeliveryCosts');
          deliverySheet.appendRow(['Sale ID', 'Date', 'Delivery Fee']);
        } else if (deliverySheet.getLastRow() === 0) {
          deliverySheet.appendRow(['Sale ID', 'Date', 'Delivery Fee']);
        }

        // Check if this sale already exists in DeliveryCosts
        const deliveryData = deliverySheet.getDataRange().getValues();
        let found = false;
        for (let i = 1; i < deliveryData.length; i++) {
          if (deliveryData[i][0] === saleId) {
            deliverySheet.getRange(i + 1, 3).setValue(sale.deliveryFee);
            found = true;
            break;
          }
        }
        if (!found) {
          deliverySheet.appendRow([saleId, sale.date, sale.deliveryFee]);
        }
      }

      // Write to SaleItems sheet
      items.forEach((item) => {
        if (hasProductNameCol) {
          itemSheet
            .getRange(itemRow, 1, 1, 7)
            .setValues([
              [
                saleId,
                item.productId,
                item.productName,
                item.quantity,
                item.isFree,
                item.unitPrice,
                item.finalPrice,
              ],
            ]);
        } else {
          itemSheet
            .getRange(itemRow, 1, 1, 6)
            .setValues([
              [
                saleId,
                item.productId,
                item.quantity,
                item.isFree,
                item.unitPrice,
                item.finalPrice,
              ],
            ]);
        }
        itemRow++;
      });
    });

    const syncedSales = saleIds.length;
    const syncedItems = itemRow - 2;

    SpreadsheetApp.getUi().alert(
      `Sync completed!\n\nSynced ${syncedSales} sales and ${syncedItems} items from Sales View to Sales and SaleItems sheets.`,
    );
  } catch (error) {
    SpreadsheetApp.getUi().alert('Error syncing data: ' + error.message);
    Logger.log('Error details: ' + error.toString());
  }
}

// =========================
// DELETE SALE FUNCTIONS
// =========================

/**
 * Get all sales for dropdown
 */
function getAllSales() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const saleSheet = ss.getSheetByName('Sales');

  if (!saleSheet || saleSheet.getLastRow() <= 1) {
    return [];
  }

  const headers = saleSheet
    .getRange(1, 1, 1, saleSheet.getLastColumn())
    .getValues()[0];
  const data = saleSheet
    .getRange(2, 1, saleSheet.getLastRow() - 1, saleSheet.getLastColumn())
    .getValues();

  const dateColIdx = 1;
  const totalColIdx = headers.indexOf('Total Amount');
  const totalCol = totalColIdx >= 0 ? totalColIdx : 2;

  return data
    .map((row) => {
      const rawDate = row[dateColIdx];
      let dateStr = '';
      if (rawDate instanceof Date) {
        dateStr = rawDate.toLocaleString();
      } else if (rawDate) {
        dateStr = String(rawDate);
      }
      return {
        id: row[0],
        date: dateStr,
        total: row[totalCol] || 0,
      };
    })
    .reverse(); // Most recent first
}

/**
 * Open sidebar to delete a sale
 */
function openDeleteSaleSidebar() {
  const template = HtmlService.createTemplateFromFile('deleteSale');
  template.sales = getAllSales();
  const html = template.evaluate().setTitle('Delete Sale').setWidth(450);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Delete a sale and restock products
 */
function deleteSale(saleId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const saleSheet = ss.getSheetByName('Sales');
    const itemSheet = ss.getSheetByName('SaleItems');
    const prodSheet = ss.getSheetByName('Products');
    const deliverySheet = ss.getSheetByName('DeliveryCosts');

    if (!saleSheet || !itemSheet || !prodSheet) {
      throw new Error('Required sheets not found');
    }

    // Get item headers to find column indices
    const itemHeaders = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductName = itemHeaders.includes('Product Name');
    const productIdColIdx = 1;
    const quantityColIdx = hasProductName ? 3 : 2;

    // Get product headers
    const prodHeaders = prodSheet
      .getRange(1, 1, 1, prodSheet.getLastColumn())
      .getValues()[0];
    const prodIdColIdx = prodHeaders.indexOf('ID');
    const prodStockColIdx = prodHeaders.indexOf('Stock');

    if (prodIdColIdx < 0 || prodStockColIdx < 0) {
      throw new Error('Product sheet missing ID or Stock column');
    }

    // Find and collect items to restock
    const itemsData = itemSheet
      .getRange(2, 1, itemSheet.getLastRow() - 1, itemSheet.getLastColumn())
      .getValues();
    const itemsToRestock = [];
    const itemRowsToDelete = [];

    for (let i = 0; i < itemsData.length; i++) {
      if (itemsData[i][0] === saleId) {
        itemsToRestock.push({
          productId: itemsData[i][productIdColIdx],
          quantity: itemsData[i][quantityColIdx] || 0,
        });
        itemRowsToDelete.push(i + 2); // +2 because row 1 is header, and array is 0-indexed
      }
    }

    // Restock products
    const prodData = prodSheet.getDataRange().getValues();
    for (const item of itemsToRestock) {
      for (let i = 1; i < prodData.length; i++) {
        if (prodData[i][prodIdColIdx] === item.productId) {
          const currentStock = prodData[i][prodStockColIdx] || 0;
          const newStock = currentStock + item.quantity;
          prodSheet.getRange(i + 1, prodStockColIdx + 1).setValue(newStock);
          break;
        }
      }
    }

    // Delete item rows (from bottom to top to avoid index shifting)
    itemRowsToDelete.sort((a, b) => b - a);
    for (const rowNum of itemRowsToDelete) {
      itemSheet.deleteRow(rowNum);
    }

    // Delete from Sales sheet
    const salesData = saleSheet.getDataRange().getValues();
    for (let i = 1; i < salesData.length; i++) {
      if (salesData[i][0] === saleId) {
        saleSheet.deleteRow(i + 1);
        break;
      }
    }

    // Delete from DeliveryCosts if exists
    if (deliverySheet && deliverySheet.getLastRow() > 1) {
      const deliveryData = deliverySheet.getDataRange().getValues();
      for (let i = deliveryData.length - 1; i >= 1; i--) {
        if (deliveryData[i][0] === saleId) {
          deliverySheet.deleteRow(i + 1);
        }
      }
    }

    // Update Sales View if exists
    const viewSheet = ss.getSheetByName('Sales View');
    if (viewSheet && viewSheet.getLastRow() > 1) {
      const viewData = viewSheet.getDataRange().getValues();
      const viewRowsToDelete = [];
      for (let i = 1; i < viewData.length; i++) {
        // Check if this row belongs to the sale (either has the sale ID or is a continuation row)
        if (viewData[i][0] === saleId) {
          viewRowsToDelete.push(i + 1);
          // Also delete continuation rows (empty sale ID, until next sale ID)
          for (let j = i + 1; j < viewData.length; j++) {
            if (viewData[j][0] === '' || viewData[j][0] === null) {
              viewRowsToDelete.push(j + 1);
            } else {
              break;
            }
          }
          break;
        }
      }
      // Delete from bottom to top
      viewRowsToDelete.sort((a, b) => b - a);
      for (const rowNum of viewRowsToDelete) {
        viewSheet.deleteRow(rowNum);
      }
    }

    return {
      success: true,
      message: `Sale ${saleId} deleted successfully!\n\nRestocked ${itemsToRestock.length} item(s).`,
    };
  } catch (error) {
    Logger.log('Error deleting sale: ' + error.toString());
    return { success: false, message: error.message };
  }
}

// =========================
// EDIT SALE FUNCTIONS
// =========================

/**
 * Get sale details for editing (includes delivery fee and discount)
 */
function getSaleDetailsForEdit(saleId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const saleSheet = ss.getSheetByName('Sales');
    const itemSheet = ss.getSheetByName('SaleItems');

    if (!saleSheet || !itemSheet) {
      return null;
    }

    // Get sale from Sales sheet
    const saleHeaders = saleSheet
      .getRange(1, 1, 1, saleSheet.getLastColumn())
      .getValues()[0];
    const saleData = saleSheet.getDataRange().getValues();

    let saleInfo = null;
    for (let i = 1; i < saleData.length; i++) {
      if (saleData[i][0] === saleId) {
        const dateColIdx = 1;
        const subtotalColIdx = saleHeaders.indexOf('Subtotal');
        const deliveryFeeColIdx = saleHeaders.indexOf('Delivery Fee');
        const discountColIdx = saleHeaders.indexOf('Discount');
        const totalColIdx = saleHeaders.indexOf('Total Amount');

        saleInfo = {
          saleId: saleId,
          date: saleData[i][dateColIdx],
          subtotal: subtotalColIdx >= 0 ? saleData[i][subtotalColIdx] || 0 : 0,
          deliveryFee:
            deliveryFeeColIdx >= 0 ? saleData[i][deliveryFeeColIdx] || 0 : 0,
          discount: discountColIdx >= 0 ? saleData[i][discountColIdx] || 0 : 0,
          totalAmount:
            totalColIdx >= 0
              ? saleData[i][totalColIdx] || 0
              : saleData[i][2] || 0,
        };
        break;
      }
    }

    if (!saleInfo) {
      return null;
    }

    // Get items from SaleItems sheet
    const itemHeaders = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductName = itemHeaders.includes('Product Name');
    const itemData = itemSheet.getDataRange().getValues();

    const productIdColIdx = 1;
    const productNameColIdx = hasProductName ? 2 : -1;
    const quantityColIdx = hasProductName ? 3 : 2;
    const isFreeColIdx = hasProductName ? 4 : 3;
    const unitPriceColIdx = hasProductName ? 5 : 4;
    const finalPriceColIdx = hasProductName ? 6 : 5;

    const items = [];
    for (let i = 1; i < itemData.length; i++) {
      if (itemData[i][0] === saleId) {
        items.push({
          productId: itemData[i][productIdColIdx],
          productName: hasProductName ? itemData[i][productNameColIdx] : '',
          quantity: itemData[i][quantityColIdx] || 0,
          isFree: itemData[i][isFreeColIdx] || false,
          unitPrice: itemData[i][unitPriceColIdx] || 0,
          finalPrice: itemData[i][finalPriceColIdx] || 0,
        });
      }
    }

    saleInfo.items = items;
    return saleInfo;
  } catch (error) {
    Logger.log('Error getting sale details for edit: ' + error.toString());
    return null;
  }
}

/**
 * Open sidebar to edit a sale
 */
function openEditSaleSidebar() {
  const template = HtmlService.createTemplateFromFile('editSale');
  template.sales = getAllSales();
  template.products = getProducts();
  const html = template.evaluate().setTitle('Edit Sale').setWidth(550);
  SpreadsheetApp.getUi().showSidebar(html);
}

/**
 * Update a sale
 */
function updateSale(saleId, saleItems, deliveryFee, saleDiscount) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    deliveryFee = deliveryFee || 0;
    saleDiscount = saleDiscount || 0;

    const saleSheet = ss.getSheetByName('Sales');
    const itemSheet = ss.getSheetByName('SaleItems');
    const prodSheet = ss.getSheetByName('Products');

    if (!saleSheet || !itemSheet || !prodSheet) {
      throw new Error('Required sheets not found');
    }

    // Get original sale items to calculate stock adjustments
    const originalItems = [];
    const itemHeaders = itemSheet
      .getRange(1, 1, 1, itemSheet.getLastColumn())
      .getValues()[0];
    const hasProductName = itemHeaders.includes('Product Name');
    const productIdColIdx = 1;
    const quantityColIdx = hasProductName ? 3 : 2;

    const allItemData = itemSheet.getDataRange().getValues();
    for (let i = 1; i < allItemData.length; i++) {
      if (allItemData[i][0] === saleId) {
        originalItems.push({
          productId: allItemData[i][productIdColIdx],
          quantity: allItemData[i][quantityColIdx] || 0,
        });
      }
    }

    // Calculate stock adjustments
    const stockAdjustments = {};
    originalItems.forEach((item) => {
      stockAdjustments[item.productId] =
        (stockAdjustments[item.productId] || 0) + item.quantity;
    });
    saleItems.forEach((item) => {
      stockAdjustments[item.productId] =
        (stockAdjustments[item.productId] || 0) - item.qty;
    });

    // Apply stock adjustments
    const prodHeaders = prodSheet
      .getRange(1, 1, 1, prodSheet.getLastColumn())
      .getValues()[0];
    const prodIdColIdx = prodHeaders.indexOf('ID');
    const prodStockColIdx = prodHeaders.indexOf('Stock');

    const prodData = prodSheet.getDataRange().getValues();
    for (const productId in stockAdjustments) {
      const adjustment = stockAdjustments[productId];
      if (adjustment !== 0) {
        for (let i = 1; i < prodData.length; i++) {
          if (prodData[i][prodIdColIdx] === productId) {
            const currentStock = prodData[i][prodStockColIdx] || 0;
            const newStock = currentStock + adjustment;
            if (newStock < 0) {
              throw new Error(
                `Insufficient stock for product. Adjustment would result in negative stock.`,
              );
            }
            prodSheet.getRange(i + 1, prodStockColIdx + 1).setValue(newStock);
            break;
          }
        }
      }
    }

    // Delete old items
    const itemRowsToDelete = [];
    for (let i = allItemData.length - 1; i >= 1; i--) {
      if (allItemData[i][0] === saleId) {
        itemRowsToDelete.push(i + 1);
      }
    }
    itemRowsToDelete.forEach((rowNum) => {
      itemSheet.deleteRow(rowNum);
    });

    // Add new items
    const products = getProducts();
    saleItems.forEach((item) => {
      const prod = products.find((p) => p.id === item.productId);
      if (!prod) {
        throw new Error(`Product with ID ${item.productId} not found`);
      }

      // Calculate unit price
      let unitPrice;
      if (item.isFree) {
        unitPrice = 0;
      } else if (
        item.unitPrice !== null &&
        item.unitPrice !== undefined &&
        item.unitPrice !== ''
      ) {
        unitPrice = Number(item.unitPrice);
      } else {
        unitPrice = prod.discount > 0 ? prod.discount : prod.price;
      }
      const finalPrice = unitPrice * item.qty;

      // Insert item
      if (hasProductName) {
        itemSheet.appendRow([
          saleId,
          item.productId,
          prod.name,
          item.qty,
          item.isFree,
          unitPrice,
          finalPrice,
        ]);
      } else {
        itemSheet.appendRow([
          saleId,
          item.productId,
          item.qty,
          item.isFree,
          unitPrice,
          finalPrice,
        ]);
      }
    });

    // Update Sales sheet
    const saleHeaders = saleSheet
      .getRange(1, 1, 1, saleSheet.getLastColumn())
      .getValues()[0];
    const saleData = saleSheet.getDataRange().getValues();

    // Calculate subtotal
    let subtotal = 0;
    saleItems.forEach((item) => {
      if (!item.isFree) {
        const prod = products.find((p) => p.id === item.productId);
        const unitPrice =
          item.unitPrice !== null &&
          item.unitPrice !== undefined &&
          item.unitPrice !== ''
            ? Number(item.unitPrice)
            : prod.discount > 0
            ? prod.discount
            : prod.price;
        subtotal += unitPrice * item.qty;
      }
    });

    const finalTotal = subtotal - deliveryFee - saleDiscount;

    // Find and update sale row
    for (let i = 1; i < saleData.length; i++) {
      if (saleData[i][0] === saleId) {
        const hasNewFormat = saleHeaders.includes('Subtotal');
        if (hasNewFormat) {
          const subtotalCol = saleHeaders.indexOf('Subtotal') + 1;
          const deliveryFeeCol = saleHeaders.indexOf('Delivery Fee') + 1;
          const discountCol = saleHeaders.indexOf('Discount') + 1;
          const totalCol = saleHeaders.indexOf('Total Amount') + 1;

          if (subtotalCol > 0)
            saleSheet.getRange(i + 1, subtotalCol).setValue(subtotal);
          if (deliveryFeeCol > 0)
            saleSheet.getRange(i + 1, deliveryFeeCol).setValue(deliveryFee);
          if (discountCol > 0)
            saleSheet.getRange(i + 1, discountCol).setValue(saleDiscount);
          if (totalCol > 0)
            saleSheet.getRange(i + 1, totalCol).setValue(finalTotal);
        } else {
          saleSheet.getRange(i + 1, 3).setValue(finalTotal);
        }
        break;
      }
    }

    // Update DeliveryCosts sheet
    let deliverySheet = ss.getSheetByName('DeliveryCosts');
    if (deliveryFee > 0) {
      if (!deliverySheet) {
        deliverySheet = ss.insertSheet('DeliveryCosts');
        deliverySheet.appendRow(['Sale ID', 'Date', 'Delivery Fee']);
      } else if (deliverySheet.getLastRow() === 0) {
        deliverySheet.appendRow(['Sale ID', 'Date', 'Delivery Fee']);
      }

      const deliveryData = deliverySheet.getDataRange().getValues();
      let found = false;
      for (let i = 1; i < deliveryData.length; i++) {
        if (deliveryData[i][0] === saleId) {
          deliverySheet.getRange(i + 1, 3).setValue(deliveryFee);
          found = true;
          break;
        }
      }
      if (!found) {
        const saleRow = saleSheet
          .getDataRange()
          .getValues()
          .find((row) => row[0] === saleId);
        if (saleRow) {
          deliverySheet.appendRow([saleId, saleRow[1], deliveryFee]);
        }
      }
    } else if (deliverySheet) {
      // Remove delivery fee entry if fee is now 0
      const deliveryData = deliverySheet.getDataRange().getValues();
      for (let i = deliveryData.length - 1; i >= 1; i--) {
        if (deliveryData[i][0] === saleId) {
          deliverySheet.deleteRow(i + 1);
          break;
        }
      }
    }

    // Update Sales View if exists
    const viewSheet = ss.getSheetByName('Sales View');
    if (viewSheet && viewSheet.getLastRow() > 1) {
      // Delete old rows for this sale
      const viewData = viewSheet.getDataRange().getValues();
      const viewRowsToDelete = [];
      for (let i = 1; i < viewData.length; i++) {
        if (viewData[i][0] === saleId) {
          viewRowsToDelete.push(i + 1);
          for (let j = i + 1; j < viewData.length; j++) {
            if (viewData[j][0] === '' || viewData[j][0] === null) {
              viewRowsToDelete.push(j + 1);
            } else {
              break;
            }
          }
          break;
        }
      }
      viewRowsToDelete.sort((a, b) => b - a);
      viewRowsToDelete.forEach((rowNum) => {
        viewSheet.deleteRow(rowNum);
      });

      // Re-add the sale
      const saleRow = saleSheet
        .getDataRange()
        .getValues()
        .find((row) => row[0] === saleId);
      if (saleRow) {
        const saleDate = saleRow[1];
        const formattedDate =
          saleDate instanceof Date
            ? saleDate.toLocaleString()
            : String(saleDate);

        // Insert at row 2 (newest first)
        let insertRow = 2;
        saleItems.forEach((item, index) => {
          const prod = products.find((p) => p.id === item.productId);
          const isFirstItem = index === 0;
          viewSheet.insertRowBefore(insertRow);

          if (isFirstItem) {
            viewSheet.getRange(insertRow, 1).setValue(saleId);
            viewSheet.getRange(insertRow, 2).setValue(formattedDate);
            viewSheet.getRange(insertRow, 7).setValue(finalTotal.toFixed(2));
            viewSheet.getRange(insertRow, 7).setFontWeight('bold');
            viewSheet.getRange(insertRow, 7).setBackground('#e8f0fe');
          }

          const unitPrice = item.isFree
            ? 0
            : item.unitPrice ||
              (prod.discount > 0 ? prod.discount : prod.price);
          const finalPrice = unitPrice * item.qty;

          viewSheet.getRange(insertRow, 3).setValue(prod.name);
          viewSheet.getRange(insertRow, 4).setValue(item.qty);
          viewSheet
            .getRange(insertRow, 5)
            .setValue(item.isFree ? 'FREE' : unitPrice.toFixed(2));
          viewSheet.getRange(insertRow, 6).setValue(finalPrice.toFixed(2));

          if (!item.isFree) {
            viewSheet.getRange(insertRow, 5).setNumberFormat('#,##0.00');
            viewSheet.getRange(insertRow, 6).setNumberFormat('#,##0.00');
          }

          insertRow++;
        });
      }
    }

    let message = `Sale updated successfully! Subtotal: $${subtotal.toFixed(
      2,
    )}`;
    if (deliveryFee > 0) message += `, Delivery: -$${deliveryFee.toFixed(2)}`;
    if (saleDiscount > 0) message += `, Discount: -$${saleDiscount.toFixed(2)}`;
    message += `, Total: $${finalTotal.toFixed(2)}`;

    return { success: true, message: message };
  } catch (error) {
    Logger.log('Error updating sale: ' + error.toString());
    return { success: false, message: error.message };
  }
}
