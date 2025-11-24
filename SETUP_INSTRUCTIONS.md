# Inventory Management System - Setup Instructions

## 📋 Prerequisites

- A Google Account
- Access to Google Sheets
- Access to Google Apps Script

## 🚀 Setup Steps

### Step 1: Create a New Google Sheet

1. Go to [Google Sheets](https://sheets.google.com)
2. Create a new spreadsheet
3. Name it something like "Inventory Management"

### Step 2: Create Required Sheets

Your spreadsheet needs **3 sheets** with the following structure:

#### Sheet 1: "Products"

Create a sheet named **"Products"** with these columns:

- **Column A**: ID (auto-generated)
- **Column B**: Name
- **Column C**: Description
- **Column D**: Price
- **Column E**: Discount
- **Column F**: Stock
- **Column G**: Created At
- **Column H**: Last Updated At (auto-populated)

**Header Row (Row 1):**

```
ID | Name | Description | Price | Discount | Stock | Created At | Last Updated At
```

**Note:** The "Last Updated At" column is automatically added when you create your first product or update an existing one. It tracks when the product information (including stock) was last modified.

#### Sheet 2: "Sales"

Create a sheet named **"Sales"** with these columns:

- **Column A**: Sale ID (auto-generated)
- **Column B**: Date
- **Column C**: Total Amount

**Header Row (Row 1):**

```
Sale ID | Date | Total Amount
```

#### Sheet 3: "SaleItems"

Create a sheet named **"SaleItems"** with these columns:

- **Column A**: Sale ID
- **Column B**: Product ID
- **Column C**: Product Name (auto-populated for readability)
- **Column D**: Quantity
- **Column E**: Is Free (TRUE/FALSE)
- **Column F**: Unit Price
- **Column G**: Final Price

**Header Row (Row 1):**

```
Sale ID | Product ID | Product Name | Quantity | Is Free | Unit Price | Final Price
```

**Note:** The "Product Name" column is automatically added when you create your first sale. If you have existing sales without product names, use the menu option "Update Product Names in Sales" to populate them.

**Note:** The "Report" sheet will be created automatically when you generate the monthly report.

### Step 3: Install the Google Apps Script Code

1. In your Google Sheet, click on **Extensions** → **Apps Script**
2. Delete any existing code in the editor
3. Copy and paste the entire contents of `Code.gs` into the editor
4. Click **File** → **Save** (or press Ctrl+S / Cmd+S)
5. Name your project (e.g., "Inventory Management")

### Step 4: Add HTML Files

For each HTML file, follow these steps:

1. In the Apps Script editor, click the **+** button next to "Files"
2. Select **HTML**
3. Name it exactly as shown (case-sensitive):
   - `form` (for Add Product)
   - `editProduct` (for Edit Product)
   - `deleteProduct` (for Delete Product)
   - `addSale` (for Add Sale)
4. Copy and paste the corresponding HTML content
5. Save the file

**Important:** The file names must match exactly:

- `form.html` → name it `form`
- `editProduct.html` → name it `editProduct`
- `deleteProduct.html` → name it `deleteProduct`
- `addSale.html` → name it `addSale`

### Step 5: Authorize the Script

1. In the Apps Script editor, click **Run** → Select any function (e.g., `onOpen`)
2. You'll be prompted to authorize the script
3. Click **Review Permissions**
4. Select your Google account
5. Click **Advanced** → **Go to [Project Name] (unsafe)**
6. Click **Allow** to grant permissions

### Step 6: Refresh Your Spreadsheet

1. Go back to your Google Sheet
2. Refresh the page (F5 or Cmd+R)
3. You should see a new menu item **"Inventory App"** in the menu bar

## ✅ Verification

1. Click on **Inventory App** in the menu
2. You should see these options:
   - Add Product
   - Edit Product
   - Delete Product
   - Add Sale
   - Generate Monthly Report

## 🐛 Troubleshooting

### Menu doesn't appear

- Make sure you've saved the `Code.gs` file
- Refresh the spreadsheet page
- Check that the `onOpen()` function is in your Code.gs

### "Sheet not found" errors

- Verify that your sheets are named exactly: "Products", "Sales", "SaleItems" (case-sensitive)
- Make sure each sheet has the header row in Row 1

### HTML files not loading

- Check that HTML file names match exactly (no .html extension in Apps Script)
- Verify the HTML content was copied correctly
- Make sure you're using `createHtmlOutputFromFile()` for `form.html` and `createTemplateFromFile()` for others

### Permission errors

- Re-authorize the script (Step 5)
- Make sure you're logged into the correct Google account

## 📝 Usage Tips

1. **Add Products First**: Before adding sales, make sure you have products in the Products sheet
2. **Stock Management**: Stock is automatically deducted when you add a sale
3. **Monthly Reports**: Click "Generate Monthly Report" to create a summary in the "Report" sheet
4. **Free Items**: Check the "Free Item" checkbox when adding items to a sale to set price to $0

## 🎨 Features

✅ Add Product - Create new products with name, description, price, discount, and stock  
✅ Edit Product - Update existing product information  
✅ Delete Product - Remove products from inventory  
✅ Add Sale - Record sales with multiple items, automatic stock deduction  
✅ Auto Price & Discount - Automatically uses discount price if available, otherwise regular price  
✅ Monthly Report - Generate revenue reports by month

## 📞 Support

If you encounter any issues, check:

1. Browser console for JavaScript errors (F12)
2. Apps Script execution log (View → Execution log)
3. Sheet names and column structure match the requirements
