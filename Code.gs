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
    .addItem('Add Product', 'openAddProductSidebar')
    .addItem('Edit Product', 'openEditProductSidebar')
    .addItem('Delete Product', 'openDeleteProductSidebar')
    .addSeparator()
    .addItem('Add Sale', 'openAddSaleSidebar')
    .addItem('View Sale Details', 'openViewSaleSidebar')
    .addSeparator()
    .addItem('Generate Monthly Report', 'generateMonthlyReport')
    .addItem('Create Combined Sales View', 'createCombinedSalesView')
    .addSeparator()
    .addItem('Update Product Names in Sales', 'updateSaleItemsWithProductNames')
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
  return rows.map((r) => ({
    id: r[idCol],
    name: r[nameCol],
    desc: r[descCol],
    price: r[priceCol],
    discount: r[discountCol],
    stock: r[stockCol],
  }));
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
function submitSale(saleItems) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create sheets
    let saleSheet = ss.getSheetByName('Sales');
    if (!saleSheet) {
      saleSheet = ss.insertSheet('Sales');
      saleSheet.appendRow(['Sale ID', 'Date', 'Total Amount']);
    } else if (saleSheet.getLastRow() === 0) {
      saleSheet.appendRow(['Sale ID', 'Date', 'Total Amount']);
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
      } else if (item.unitPrice !== null && item.unitPrice !== undefined && item.unitPrice !== '') {
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

    // Insert sale summary
    saleSheet.appendRow([saleId, date, totalAmount]);

    // Automatically append new sale to Sales View sheet if it exists
    try {
      const viewSheet = ss.getSheetByName('Sales View');
      if (viewSheet && viewSheet.getLastRow() > 0) {
        // Only update if the sheet exists and has data
        // Use a small delay to ensure the data is written before reading
        Utilities.sleep(100);
        appendNewSaleToView(saleId, date, totalAmount, saleItems);
      }
    } catch (viewError) {
      // Silently fail if view update fails - don't break the sale submission
      Logger.log('Error updating Sales View: ' + viewError.toString());
    }

    return {
      success: true,
      message: `Sale recorded successfully! Total: $${totalAmount.toFixed(2)}`,
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

    // Get all sales with dates
    const salesData = saleSheet
      .getRange(2, 1, saleSheet.getLastRow() - 1, 3)
      .getValues();

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
    const quantityColIndex = hasProductName ? 3 : 2; // Quantity column
    const finalPriceColIndex = hasProductName ? 6 : 5; // Final Price column

    // Create map of saleId to date and total amount
    const saleDateMap = {};
    const saleAmountMap = {};
    salesData.forEach((row) => {
      saleDateMap[row[0]] = row[1]; // saleId -> date
      saleAmountMap[row[0]] = row[2]; // saleId -> total amount
    });

    // Aggregate data by month
    const monthDataMap = {};

    // Process items for revenue and quantity
    itemsData.forEach((row) => {
      const saleId = row[0];
      const quantity = row[quantityColIndex] || 0;
      const finalPrice = row[finalPriceColIndex] || 0;
      const date = saleDateMap[saleId];

      if (date) {
        const year = date.getFullYear();
        const month = date.getMonth() + 1;
        const key = `${year}-${String(month).padStart(2, '0')}`;

        if (!monthDataMap[key]) {
          monthDataMap[key] = {
            revenue: 0,
            itemsSold: 0,
            transactions: new Set(),
          };
        }
        monthDataMap[key].revenue += finalPrice;
        monthDataMap[key].itemsSold += quantity;
        monthDataMap[key].transactions.add(saleId);
      }
    });

    // Sort by year-month
    const sortedKeys = Object.keys(monthDataMap).sort();

    // Write header
    reportSheet.appendRow([
      'Year',
      'Month',
      'Revenue',
      'Transactions',
      'Items Sold',
      'Avg Sale',
      'Growth %',
    ]);

    // Write monthly data
    let previousRevenue = null;
    sortedKeys.forEach((key) => {
      const [year, month] = key.split('-');
      const data = monthDataMap[key];
      const revenue = data.revenue;
      const transactions = data.transactions.size;
      const itemsSold = data.itemsSold;
      const avgSale = transactions > 0 ? revenue / transactions : 0;

      // Calculate growth percentage
      let growth = '';
      if (previousRevenue !== null && previousRevenue > 0) {
        const growthPercent =
          ((revenue - previousRevenue) / previousRevenue) * 100;
        growth =
          growthPercent >= 0
            ? `+${growthPercent.toFixed(1)}%`
            : `${growthPercent.toFixed(1)}%`;
      } else {
        growth = 'N/A';
      }

      reportSheet.appendRow([
        year,
        month,
        revenue.toFixed(2),
        transactions,
        itemsSold,
        avgSale.toFixed(2),
        growth,
      ]);

      previousRevenue = revenue;
    });

    // Add summary section
    const summaryRow = sortedKeys.length + 3;
    reportSheet.getRange(summaryRow, 1).setValue('SUMMARY');
    reportSheet.getRange(summaryRow, 1).setFontWeight('bold');
    reportSheet.getRange(summaryRow, 1).setFontSize(12);

    // Calculate totals
    let totalRevenue = 0;
    let totalTransactions = 0;
    let totalItemsSold = 0;

    sortedKeys.forEach((key) => {
      const data = monthDataMap[key];
      totalRevenue += data.revenue;
      totalTransactions += data.transactions.size;
      totalItemsSold += data.itemsSold;
    });

    const overallAvgSale =
      totalTransactions > 0 ? totalRevenue / totalTransactions : 0;

    reportSheet.appendRow(['Total Revenue:', totalRevenue.toFixed(2)]);
    reportSheet.appendRow(['Total Transactions:', totalTransactions]);
    reportSheet.appendRow(['Total Items Sold:', totalItemsSold]);
    reportSheet.appendRow(['Overall Avg Sale:', overallAvgSale.toFixed(2)]);
    reportSheet.appendRow(['Periods:', sortedKeys.length + ' months']);

    // Format the report
    const headerRange = reportSheet.getRange(1, 1, 1, 7);
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285f4');
    headerRange.setFontColor('#ffffff');

    // Set column widths
    reportSheet.setColumnWidth(1, 60); // Year
    reportSheet.setColumnWidth(2, 60); // Month
    reportSheet.setColumnWidth(3, 100); // Revenue
    reportSheet.setColumnWidth(4, 100); // Transactions
    reportSheet.setColumnWidth(5, 100); // Items Sold
    reportSheet.setColumnWidth(6, 100); // Avg Sale
    reportSheet.setColumnWidth(7, 90); // Growth %

    // Format summary section
    const summaryStartRow = summaryRow;
    const summaryEndRow = summaryRow + 5;
    reportSheet
      .getRange(summaryStartRow, 1, summaryEndRow - summaryStartRow + 1, 1)
      .setFontWeight('bold');
    reportSheet
      .getRange(summaryStartRow, 2, summaryEndRow - summaryStartRow + 1, 1)
      .setNumberFormat('#,##0.00');

    // Add borders
    const dataRange = reportSheet.getRange(1, 1, sortedKeys.length + 1, 7);
    dataRange.setBorder(true, true, true, true, true, true);

    SpreadsheetApp.getUi().alert(
      "Comprehensive monthly report generated successfully in the 'Report' sheet!",
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
