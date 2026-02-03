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
 * Pricing tier is determined per product+color combination.
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
    summarySheet.clearFormats();
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

  // Count items by product+color (this is the key unit for pricing tiers)
  const productColorData = {}; // { "Better Sweater Jacket|Black": { count: 5, embroideryCount: 2 } }

  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var product = row[productIdx];
    var color = normalizeColor(row[colorIdx]);
    var embName = row[embroideredNameIdx];

    if (!product) continue;

    var key = product + "|" + color;
    if (!productColorData[key]) {
      productColorData[key] = { product: product, color: color, count: 0, embroideryCount: 0 };
    }
    productColorData[key].count++;
    if (embName && embName.toString().trim()) {
      productColorData[key].embroideryCount++;
    }
  }

  // Calculate totals
  var grandTotalItems = 0;
  var grandTotalCost = 0;
  var grandTotalEmbroidery = 0;
  var hasUnfulfilled = false;

  // Process each product+color combo
  var comboKeys = Object.keys(productColorData).sort();
  var comboRows = [];

  for (var k = 0; k < comboKeys.length; k++) {
    var combo = productColorData[comboKeys[k]];
    var tier = getPriceTier(combo.count);
    var unitPrice = PRICING[combo.product] ? PRICING[combo.product][tier] : 0;
    var subtotal = combo.count * unitPrice;
    var embFees = combo.embroideryCount * EMBROIDERY_FEE;

    var status = "";
    if (combo.count < 6) {
      status = "⚠️ NOT FULFILLED";
      hasUnfulfilled = true;
    } else {
      status = "✓ OK";
      grandTotalItems += combo.count;
      grandTotalCost += subtotal;
      grandTotalEmbroidery += embFees;
    }

    comboRows.push({
      product: combo.product,
      color: combo.color,
      count: combo.count,
      tier: tier,
      unitPrice: unitPrice,
      subtotal: subtotal,
      embroideryCount: combo.embroideryCount,
      embFees: embFees,
      status: status
    });
  }

  // Build summary output
  var output = [];
  var rowTracker = {}; // Track special rows for formatting

  // Header section
  output.push(["ORDER SUMMARY", "", "", "", "", "", ""]);
  rowTracker.title = output.length;

  output.push(["", "", "", "", "", "", ""]);

  // Main breakdown table
  output.push(["PRODUCT + COLOR BREAKDOWN (Tier is per combo)", "", "", "", "", "", ""]);
  rowTracker.tableTitle = output.length;

  output.push(["Product", "Color", "Qty", "Tier", "Unit Price", "Subtotal", "Status"]);
  rowTracker.tableHeader = output.length;

  var currentProduct = "";
  for (var j = 0; j < comboRows.length; j++) {
    var cr = comboRows[j];

    // Add blank row between products for readability
    if (cr.product !== currentProduct && currentProduct !== "") {
      output.push(["", "", "", "", "", "", ""]);
    }
    currentProduct = cr.product;

    output.push([
      cr.product,
      cr.color,
      String(cr.count),
      cr.count < 6 ? "N/A" : cr.tier + "+",
      cr.count < 6 ? "-" : "$" + cr.unitPrice.toFixed(2),
      cr.count < 6 ? "-" : "$" + cr.subtotal.toFixed(2),
      cr.status
    ]);
  }

  output.push(["", "", "", "", "", "", ""]);

  // Totals (only for fulfilled items)
  output.push(["TOTALS (fulfilled items only)", "", "", "", "", "", ""]);
  rowTracker.totalsTitle = output.length;

  output.push(["Total Items:", String(grandTotalItems), "", "Product Cost:", "$" + grandTotalCost.toFixed(2), "", ""]);
  output.push(["", "", "", "Embroidery:", "$" + grandTotalEmbroidery.toFixed(2), "", ""]);
  output.push(["", "", "", "GRAND TOTAL:", "$" + (grandTotalCost + grandTotalEmbroidery).toFixed(2), "", ""]);
  rowTracker.grandTotal = output.length;

  // Warning if any unfulfilled
  if (hasUnfulfilled) {
    output.push(["", "", "", "", "", "", ""]);
    output.push(["⚠️ WARNING: Combos marked 'NOT FULFILLED' are below the 6-item minimum and will not be ordered.", "", "", "", "", "", ""]);
    rowTracker.warning = output.length;
  }

  // Write to sheet
  summarySheet.getRange(1, 1, output.length, 7).setValues(output);

  // Apply formatting
  summarySheet.getRange(rowTracker.title, 1).setFontWeight("bold").setFontSize(14);
  summarySheet.getRange(rowTracker.tableTitle, 1).setFontWeight("bold").setFontSize(11);
  summarySheet.getRange(rowTracker.tableHeader, 1, 1, 7).setFontWeight("bold").setBackground("#e8e2da");
  summarySheet.getRange(rowTracker.totalsTitle, 1).setFontWeight("bold").setFontSize(11);
  summarySheet.getRange(rowTracker.grandTotal, 4, 1, 2).setFontWeight("bold").setBackground("#d4f5d4");

  if (rowTracker.warning) {
    summarySheet.getRange(rowTracker.warning, 1).setFontWeight("bold").setFontColor("#b5403a");
  }

  // Auto-resize columns
  summarySheet.autoResizeColumns(1, 7);

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
