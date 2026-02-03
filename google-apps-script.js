/**
 * Google Apps Script — deploy as a Web App to receive form submissions.
 *
 * Setup:
 * 1. Create a new Google Sheet
 * 2. Go to Extensions > Apps Script
 * 3. Paste this entire file into Code.gs (replace any existing code)
 * 4. Click Deploy > New Deployment
 * 5. Select "Web app"
 * 6. Set "Execute as" = Me, "Who has access" = Anyone
 * 7. Copy the deployment URL and paste it into index.html (APPS_SCRIPT_URL)
 */

// ─── PRICING CONFIG ───────────────────────────────────────────────────────────
// Matches pricing.json — update here if prices change
const PRICING = {
  "Better Sweater Jacket": { 6: 175.00, 18: 173.58, 50: 162.68, 72: 159.68 },
  "Better Sweater Vest": { 6: 132.88, 18: 131.28, 50: 123.58, 72: 119.88 },
  "Better Sweater Quarter Zip": { 6: 155.00, 18: 152.00, 50: 142.68, 72: 139.58 },
};
const EMBROIDERY_FEE = 8.00;

// Colors that count as "Gray" for grouping purposes
const GRAY_COLORS = ["Birch White", "Stonewash"];

// ──────────────────────────────────────────────────────────────────────────────

/**
 * Normalizes color for grouping: Birch White and Stonewash → "Gray"
 */
function normalizeColor(color) {
  if (GRAY_COLORS.indexOf(color) !== -1) {
    return "Gray";
  }
  return color;
}

/**
 * Get the price tier based on total quantity of a product
 */
function getPriceTier(qty) {
  if (qty >= 72) return 72;
  if (qty >= 50) return 50;
  if (qty >= 18) return 18;
  return 6;
}

/**
 * Creates or updates the Summary tab with item counts, pricing, and totals.
 * Can be run manually from the Apps Script editor or via custom menu.
 */
function updateSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ordersSheet = ss.getSheetByName("Orders") || ss.getSheets()[0];

  // Get or create Summary sheet
  let summarySheet = ss.getSheetByName("Summary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("Summary");
  } else {
    summarySheet.clear();
  }

  // Read order data
  const data = ordersSheet.getDataRange().getValues();
  if (data.length <= 1) {
    summarySheet.getRange("A1").setValue("No orders yet.");
    return;
  }

  const headers = data[0];
  const productIdx = headers.indexOf("Product");
  const colorIdx = headers.indexOf("Color");
  const embroideredNameIdx = headers.indexOf("Embroidered Name");

  // Count items by product and by product+color
  const productCounts = {};      // { "Better Sweater Jacket": 10 }
  const productColorCounts = {}; // { "Better Sweater Jacket|Black": 5 }
  const embroideryCount = {};    // { "Better Sweater Jacket": 3 } — items with embroidered names

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var product = row[productIdx];
    var color = normalizeColor(row[colorIdx]);
    var embName = row[embroideredNameIdx];

    if (!product) continue;

    // Product totals
    productCounts[product] = (productCounts[product] || 0) + 1;

    // Product + Color breakdown
    var key = product + "|" + color;
    productColorCounts[key] = (productColorCounts[key] || 0) + 1;

    // Embroidery count
    if (embName && embName.toString().trim()) {
      embroideryCount[product] = (embroideryCount[product] || 0) + 1;
    }
  }

  // Calculate grand total for tier determination
  var grandTotal = 0;
  for (var p in productCounts) {
    grandTotal += productCounts[p];
  }
  var tier = getPriceTier(grandTotal);

  // Build summary output
  var output = [];

  // Header section
  output.push(["ORDER SUMMARY", "", "", "", ""]);
  output.push(["Total Items:", grandTotal, "", "Price Tier:", tier + "+ pcs"]);
  output.push(["", "", "", "", ""]);

  // Product breakdown header
  output.push(["PRODUCT BREAKDOWN", "", "", "", ""]);
  output.push(["Product", "Qty", "Unit Price", "Subtotal", "Embroidery Fees"]);

  var productSubtotal = 0;
  var embroiderySubtotal = 0;
  var products = Object.keys(productCounts).sort();

  for (var j = 0; j < products.length; j++) {
    var prod = products[j];
    var qty = productCounts[prod];
    var unitPrice = PRICING[prod] ? PRICING[prod][tier] : 0;
    var subtotal = qty * unitPrice;
    var embQty = embroideryCount[prod] || 0;
    var embFees = embQty * EMBROIDERY_FEE;

    output.push([prod, qty, "$" + unitPrice.toFixed(2), "$" + subtotal.toFixed(2), embQty > 0 ? embQty + " × $" + EMBROIDERY_FEE.toFixed(2) + " = $" + embFees.toFixed(2) : "-"]);

    productSubtotal += subtotal;
    embroiderySubtotal += embFees;
  }

  output.push(["", "", "", "", ""]);
  output.push(["", "", "Product Total:", "$" + productSubtotal.toFixed(2), ""]);
  output.push(["", "", "Embroidery Total:", "$" + embroiderySubtotal.toFixed(2), ""]);
  output.push(["", "", "GRAND TOTAL:", "$" + (productSubtotal + embroiderySubtotal).toFixed(2), ""]);
  output.push(["", "", "", "", ""]);

  // Color breakdown header
  output.push(["COLOR BREAKDOWN BY PRODUCT", "", "", "", ""]);
  output.push(["Product", "Color", "Qty", "", ""]);

  // Sort keys and output
  var colorKeys = Object.keys(productColorCounts).sort();
  var currentProduct = "";
  for (var k = 0; k < colorKeys.length; k++) {
    var parts = colorKeys[k].split("|");
    var prodName = parts[0];
    var colorName = parts[1];
    var count = productColorCounts[colorKeys[k]];

    // Add blank row between products for readability
    if (prodName !== currentProduct && currentProduct !== "") {
      output.push(["", "", "", "", ""]);
    }
    currentProduct = prodName;

    output.push([prodName, colorName, count, "", ""]);
  }

  // Write to sheet
  summarySheet.getRange(1, 1, output.length, 5).setValues(output);

  // Format headers
  summarySheet.getRange("A1").setFontWeight("bold").setFontSize(14);
  summarySheet.getRange("A4").setFontWeight("bold").setFontSize(12);
  summarySheet.getRange("A5:E5").setFontWeight("bold").setBackground("#e8e2da");
  summarySheet.getRange("A" + (output.length - colorKeys.length - 1)).setFontWeight("bold").setFontSize(12);
  summarySheet.getRange("A" + (output.length - colorKeys.length) + ":E" + (output.length - colorKeys.length)).setFontWeight("bold").setBackground("#e8e2da");

  // Grand total formatting
  for (var r = 1; r <= output.length; r++) {
    if (output[r-1][2] === "GRAND TOTAL:") {
      summarySheet.getRange(r, 3, 1, 2).setFontWeight("bold").setBackground("#d4f5d4");
    }
  }

  // Auto-resize columns
  summarySheet.autoResizeColumns(1, 5);

  SpreadsheetApp.getActiveSpreadsheet().toast("Summary updated!", "Success");
}

/**
 * Adds a custom menu to easily refresh the summary
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Order Tools")
    .addItem("Update Summary", "updateSummary")
    .addToUi();
}

/**
 * GET endpoint — returns all orders as JSON for the invoice script.
 * Usage: fetch(APPS_SCRIPT_URL) returns { orders: [...] }
 */
function doGet(e) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName("Orders") || ss.getSheets()[0];
    const data = sheet.getDataRange().getValues();

    if (data.length <= 1) {
      return ContentService.createTextOutput(
        JSON.stringify({ orders: [] })
      ).setMimeType(ContentService.MimeType.JSON);
    }

    const headers = data[0];
    const orders = data.slice(1).map(function(row) {
      const order = {};
      headers.forEach(function(header, i) {
        order[header] = row[i];
      });
      return order;
    });

    return ContentService.createTextOutput(
      JSON.stringify({ orders: orders })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Get or create Orders sheet
    let sheet = ss.getSheetByName("Orders");
    if (!sheet) {
      // Rename first sheet to "Orders" if it exists, otherwise create it
      sheet = ss.getSheets()[0];
      if (sheet) {
        sheet.setName("Orders");
      } else {
        sheet = ss.insertSheet("Orders");
      }
    }

    // Write header row if sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Name",
        "Phone",
        "Email",
        "Position",
        "Product",
        "Style",
        "Size",
        "Color",
        "Logo",
        "Embroidered Name",
        "Thread Color",
      ]);
    }

    const timestamp = new Date().toISOString();

    // One row per line item, person info repeated
    data.items.forEach(function (item) {
      sheet.appendRow([
        timestamp,
        data.name,
        data.phone,
        data.email,
        data.position,
        item.product,
        item.style,
        item.size,
        item.color,
        item.logo,
        item.embroideredName || "",
        item.threadColor || "",
      ]);
    });

    return ContentService.createTextOutput(
      JSON.stringify({ status: "ok" })
    ).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(
      JSON.stringify({ status: "error", message: err.toString() })
    ).setMimeType(ContentService.MimeType.JSON);
  }
}
