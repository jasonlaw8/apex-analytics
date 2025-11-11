/**
 * ========================================
 * APEX GOLF ANALYTICS - FILE 4 OF 4
 * ========================================
 * 
 * THIS FILE CONTAINS:
 * - Helper/utility functions (normalization, calculation, quartiles)
 * - Category override system
 * - Transaction override system
 * - Booking payment validation
 * - All supporting functions that rarely change
 * 
 * This file is the most stable - rarely needs updates
 */

// ============================================
// HELPER FUNCTIONS - NORMALIZATION
// ============================================

function normalizeEmail(email) {
  if (!email || email === "") return null;
  return String(email).toLowerCase().trim();
}

function normalizePhone(phone) {
  if (!phone || phone === "") return null;
  var phoneStr = String(phone).replace(/\D/g, '');
  if (phoneStr === "") return null;
  return phoneStr.length > 10 ? phoneStr.slice(-10) : phoneStr;
}

function normalizeString(str) {
  if (!str || str === "") return null;
  return String(str).toLowerCase().trim();
}

// ============================================
// HELPER FUNCTIONS - CALCULATIONS
// ============================================

function calculateQuartiles(arr, prefix) {
  if (!arr || arr.length === 0) return [[prefix + "No data", ""]];
  
  var sorted = arr.slice().sort(function(a, b) { return a - b; });
  var n = sorted.length;
  
  var q0 = sorted[0];
  var q25 = sorted[Math.floor(n * 0.25)];
  var q50 = sorted[Math.floor(n * 0.50)];
  var q75 = sorted[Math.floor(n * 0.75)];
  var q100 = sorted[n - 1];
  
  return [
    ["  Min (0%)", prefix + q0.toFixed(2)],
    ["  25th percentile", prefix + q25.toFixed(2)],
    ["  Median (50%)", prefix + q50.toFixed(2)],
    ["  75th percentile", prefix + q75.toFixed(2)],
    ["  Max (100%)", prefix + q100.toFixed(2)]
  ];
}

function calculateMedian(arr) {
  if (arr.length === 0) return 0;
  var sorted = arr.slice().sort(function(a, b) { return a - b; });
  var mid = Math.floor(sorted.length / 2);
  return sorted.length % 2 === 0 ? (sorted[mid - 1] + sorted[mid]) / 2 : sorted[mid];
}

function calculateAverageTimeBetweenVisits(transData, customerIdCol, dateCol) {
  var customerVisits = {};
  
  for (var i = 1; i < transData.length; i++) {
    var customerId = transData[i][customerIdCol];
    var date = transData[i][dateCol];
    
    if (!customerId || !date) continue;
    
    if (!customerVisits[customerId]) {
      customerVisits[customerId] = [];
    }
    customerVisits[customerId].push(date);
  }
  
  var timeDiffs = [];
  for (var customerId in customerVisits) {
    if (customerVisits[customerId].length > 1) {
      var visits = customerVisits[customerId].sort(function(a, b) { return a - b; });
      var daysDiff = (visits[1] - visits[0]) / (1000 * 60 * 60 * 24);
      timeDiffs.push(daysDiff);
    }
  }
  
  if (timeDiffs.length === 0) return "N/A";
  var avg = timeDiffs.reduce(function(a, b) { return a + b; }, 0) / timeDiffs.length;
  return avg.toFixed(1) + " days";
}

function calculateRepeatPurchaseRates(transData, customerIdCol, dateCol) {
  var customerPurchases = {};
  
  for (var i = 1; i < transData.length; i++) {
    var customerId = transData[i][customerIdCol];
    var date = transData[i][dateCol];
    
    if (!customerId || !date) continue;
    
    if (!customerPurchases[customerId]) {
      customerPurchases[customerId] = [];
    }
    customerPurchases[customerId].push(date);
  }
  
  var repeat30 = 0, repeat60 = 0, repeat90 = 0;
  var totalCustomers = Object.keys(customerPurchases).length;
  
  for (var customerId in customerPurchases) {
    if (customerPurchases[customerId].length > 1) {
      var purchases = customerPurchases[customerId].sort(function(a, b) { return a - b; });
      var daysBetween = (purchases[1] - purchases[0]) / (1000 * 60 * 60 * 24);
      
      if (daysBetween <= 30) repeat30++;
      if (daysBetween <= 60) repeat60++;
      if (daysBetween <= 90) repeat90++;
    }
  }
  
  return {
    thirtyDays: (repeat30 / totalCustomers * 100).toFixed(1),
    sixtyDays: (repeat60 / totalCustomers * 100).toFixed(1),
    ninetyDays: (repeat90 / totalCustomers * 100).toFixed(1)
  };
}

// ============================================
// CATEGORY SYSTEM
// ============================================

/**
 * Maps detailed categories to major categories
 * Caches Category Override sheet to avoid repeated reads
 * Priority: Item Name > Category > Default Logic
 */
function getMajorCategory(category, itemName) {
  if (!category) return "Miscellaneous";
  
  var cat = String(category).toLowerCase().trim();
  var item = itemName ? String(itemName).toLowerCase().trim() : null;
  
  // Check cache first - create cache key from both category and item
  var cache = CacheService.getScriptCache();
  var cacheKey = "majorCat_" + cat + "_" + (item || "");
  var cached = cache.get(cacheKey);
  
  if (cached !== null) {
    return cached;
  }
  
  // Not in cache, check Category Override sheet
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var overrideSheet = ss.getSheetByName("Category Override");
  
  var result = null;
  
  if (overrideSheet) {
    var lastRow = overrideSheet.getLastRow();
    if (lastRow > 5) {  // Skip header rows
      var data = overrideSheet.getRange(6, 1, lastRow - 5, 3).getValues();
      
      for (var i = 0; i < data.length; i++) {
        var overrideCat = data[i][0];
        var overrideItem = data[i][1];
        var majorCat = data[i][2];
        
        if (!majorCat) continue;
        
        // Priority 1: Item Name match (HIGHEST)
        if (item && overrideItem && String(overrideItem).toLowerCase().trim() === item) {
          result = majorCat;
          break; // Found item match, stop looking
        }
        
        // Priority 2: Category match
        if (!result && overrideCat && String(overrideCat).toLowerCase().trim() === cat) {
          result = majorCat;
          // Don't break - keep looking for item match
        }
      }
    }
  }
  
  // If found in override, cache and return
  if (result) {
    cache.put(cacheKey, result, 360);
    return result;
  }
  
  // Not in override sheet, use default logic
  result = getDefaultMajorCategory(cat);
  
  // Cache the default result too
  cache.put(cacheKey, result, 360);
  
  return result;
}

/**
 * Default category mapping logic (when not overridden)
 */
function getDefaultMajorCategory(cat) {
  // Food categories
  if (cat === "food" || 
      cat === "handhelds" || 
      cat === "desserts" || 
      cat === "shareables" ||
      cat === "starters" ||
      cat === "pizzas" ||
      cat === "foodshareables" ||
      cat === "food shareables" ||
      cat.includes("pizza") ||
      cat.includes("dessert") ||
      cat.includes("handheld")) {
    return "Food";
  }
  
  // Beverage categories
  if (cat === "beverage" || 
      cat === "beveragenon-alcoholic" ||
      cat === "beverage non-alcoholic" ||
      cat === "beveragedrink menu" ||
      cat === "beverage drink menu" ||
      cat === "alcoholic beverage" ||
      cat === "cocktails" ||
      cat === "bar" ||
      cat === "happy hour" ||
      cat.includes("beverage") ||
      cat.includes("drink") ||
      cat.includes("cocktail") ||
      cat.includes("beer") ||
      cat.includes("wine") ||
      cat.includes("alcohol")) {
    return "Beverage";
  }
  
  // Golf categories
  if (cat === "bay rental" || 
      cat === "golf bay rental" ||
      cat === "bay rentalhandhelds" ||
      cat === "whoosh sim revenue" ||
      cat === "golf" ||
      cat.includes("bay") ||
      cat.includes("golf") ||
      cat.includes("sim")) {
    return "Golf";
  }
  
  // Membership categories
  if (cat === "membership" ||
      cat === "memberships" ||
      cat.includes("membership") ||
      cat.includes("member")) {
    return "Membership";
  }
  
  // Event categories
  if (cat === "event" ||
      cat === "events" ||
      cat.includes("event") ||
      cat.includes("private") ||
      cat.includes("corporate") ||
      cat.includes("party") ||
      cat.includes("tournament") ||
      cat.includes("league")) {
    return "Event";
  }
  
  // Everything else
  return "Miscellaneous";
}

function getCategoryOverrides() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var overrideSheet = ss.getSheetByName("Category Override");
  
  if (!overrideSheet) {
    return {categories: {}, items: {}};
  }
  
  var lastRow = overrideSheet.getLastRow();
  if (lastRow <= 5) {
    return {categories: {}, items: {}};
  }
  
  // Read columns A, B, C (Category, Item Name, Major Category)
  // Start at row 6 to skip headers/instructions
  var data = overrideSheet.getRange(6, 1, lastRow - 5, 3).getValues();
  var categoryOverrides = {};
  var itemOverrides = {};
  
  for (var i = 0; i < data.length; i++) {
    var categoryName = data[i][0];
    var itemName = data[i][1];
    var majorCategory = data[i][2];
    
    // Stop at first completely empty row
    if (!categoryName && !itemName && !majorCategory) {
      break;
    }
    
    // Skip example rows
    if (categoryName && (String(categoryName).indexOf("EXAMPLE") >= 0 || String(categoryName).indexOf("Bay Rental") >= 0)) {
      if (i < 10) continue; // Only skip if in first 10 rows (example area)
    }
    
    // Item Name override (Column B) - HIGHEST PRIORITY
    if (itemName && majorCategory) {
      itemOverrides[itemName.toLowerCase().trim()] = majorCategory;
    }
    
    // Category override (Column A)
    if (categoryName && majorCategory) {
      categoryOverrides[categoryName.toLowerCase().trim()] = majorCategory;
    }
  }
  
  return {categories: categoryOverrides, items: itemOverrides};
}

function getTransactionOverrides() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var overrideSheet = ss.getSheetByName("Item Transaction Override");
  
  if (!overrideSheet) {
    return {};
  }
  
  var lastRow = overrideSheet.getLastRow();
  if (lastRow <= 1) {
    return {};  // No data
  }
  
  // Only read actual data, not thousands of empty rows
  var data = overrideSheet.getRange(1, 1, lastRow, 3).getValues();
  var overrides = {};
  
  for (var i = 1; i < data.length; i++) {
    var transactionId = data[i][0];
    var category = data[i][2];
    
    // Stop at first completely empty row
    if (!transactionId && !category) {
      break;
    }
    
    if (transactionId && category) {
      overrides[transactionId] = category;
    }
  }
  
  return overrides;
}

// ============================================
// CATEGORY OVERRIDE SETUP
// ============================================

function setupCategoryOverride() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var overrideSheet = ss.getSheetByName("Category Override");
  
  if (!overrideSheet) {
    overrideSheet = ss.insertSheet("Category Override");
  }
  
  overrideSheet.clear();
  
  overrideSheet.getRange("A1:D1").merge();
  overrideSheet.getRange("A1").setValue("📌 CATEGORY & ITEM OVERRIDE - Permanent Category Corrections");
  overrideSheet.getRange("A1").setFontSize(14).setFontWeight("bold").setBackground("#EA4335").setFontColor("white");
  
  overrideSheet.getRange("A2:D2").merge();
  overrideSheet.getRange("A2").setValue("Use this sheet to permanently fix incorrect category mappings. Item Name overrides take priority over Category overrides.");
  overrideSheet.getRange("A2").setFontSize(10).setFontColor("#666666");
  
  overrideSheet.getRange("A3:D3").merge();
  overrideSheet.getRange("A3").setValue("Fill EITHER Category OR Item Name (or both). Item Name = specific item. Category = all items in that category.");
  overrideSheet.getRange("A3").setFontSize(9).setFontStyle("italic").setBackground("#fff3cd");
  
  var headers = ["Category (from Square)", "Item Name (from Square)", "Major Category", "Notes (optional)"];
  overrideSheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  overrideSheet.getRange(5, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8E8E8");
  
  overrideSheet.getRange("A6").setValue("Valid Major Categories:");
  overrideSheet.getRange("A6").setFontWeight("bold");
  
  var validCategories = [
    ["Food"],
    ["Beverage"],
    ["Golf"],
    ["Membership"],
    ["Miscellaneous"],
    ["Event"]
  ];
  
  overrideSheet.getRange(7, 1, validCategories.length, 1).setValues(validCategories);
  overrideSheet.getRange(7, 1, validCategories.length, 1).setBackground("#e8f5e9");
  
  overrideSheet.getRange("A14").setValue("EXAMPLE ENTRIES:");
  overrideSheet.getRange("A14").setFontWeight("bold").setBackground("#fff9c4");
  
  var examples = [
    ["Bay Rental", "", "Golf", "Override entire category"],
    ["", "Pepsi", "Beverage", "Override just this item"],
    ["Shareables", "", "Food", "All shareables are food"],
    ["", "Nachos", "Food", "Just nachos item"],
    ["Misc", "Bay 1", "Golf", "Bay 1 specifically (if also in Misc category)"]
  ];
  
  overrideSheet.getRange(15, 1, examples.length, 4).setValues(examples);
  overrideSheet.getRange(15, 1, examples.length, 4).setBackground("#f5f5f5").setFontStyle("italic");
  
  overrideSheet.setColumnWidth(1, 250);
  overrideSheet.setColumnWidth(2, 250);
  overrideSheet.setColumnWidth(3, 150);
  overrideSheet.setColumnWidth(4, 300);
  overrideSheet.setFrozenRows(5);
  
  var validationRange = overrideSheet.getRange("C15:C1000");
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Food", "Beverage", "Golf", "Membership", "Miscellaneous", "Event"], true)
    .setAllowInvalid(false)
    .setHelpText("Select a valid major category")
    .build();
  validationRange.setDataValidation(rule);
  
  ss.setActiveSheet(overrideSheet);
  
  ui.alert("✅ Category & Item Override Sheet Ready!",
    "Add rows below row 14 with:\n" +
    "Column A: Category name from Square (optional)\n" +
    "Column B: Item name from Square (optional)\n" +
    "Column C: Major category (Food/Beverage/Golf/Membership/Miscellaneous/Event)\n" +
    "Column D: Optional notes\n\n" +
    "Fill EITHER Column A OR Column B (or both).\n" +
    "Item Name (Column B) takes priority over Category (Column A).\n\n" +
    "These overrides will be applied every time you run analysis.",
    ui.ButtonSet.OK);
}

function viewCategoryOverrides() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var overrideSheet = ss.getSheetByName("Category Override");
  
  if (!overrideSheet) {
    var response = ui.alert("No Override Sheet Found",
      "Would you like to create the Category Override sheet?",
      ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.YES) {
      setupCategoryOverride();
    }
    return;
  }
  
  var overrideData = overrideSheet.getDataRange().getValues();
  var count = 0;
  var report = "ACTIVE CATEGORY OVERRIDES:\n\n";
  
  for (var i = 1; i < overrideData.length; i++) {
    var cat = overrideData[i][0];
    var majorCat = overrideData[i][1];
    var notes = overrideData[i][2];
    
    if (cat && majorCat && !String(cat).includes("EXAMPLE") && !String(cat).includes("Wrong Category")) {
      count++;
      report += count + ". \"" + cat + "\" → " + majorCat;
      if (notes) {
        report += " (" + notes + ")";
      }
      report += "\n";
    }
  }
  
  if (count === 0) {
    report += "No active overrides found.\n\n";
    report += "Add overrides to the 'Category Override' sheet to permanently fix incorrect categorizations.";
  } else {
    report += "\nTotal overrides: " + count;
  }
  
  ui.alert("Category Overrides", report, ui.ButtonSet.OK);
}

// ============================================
// TRANSACTION OVERRIDE SYSTEM
// ============================================

function setupItemTransactionOverride() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var overrideSheet = ss.getSheetByName("Item Transaction Override");
  
  if (!overrideSheet) {
    overrideSheet = ss.insertSheet("Item Transaction Override");
  }
  
  overrideSheet.clear();
  
  overrideSheet.getRange("A1:D1").merge();
  overrideSheet.getRange("A1").setValue("🔧 ITEM TRANSACTION OVERRIDE - Fix Individual Transaction Categories");
  overrideSheet.getRange("A1").setFontSize(14).setFontWeight("bold").setBackground("#673AB7").setFontColor("white");
  
  overrideSheet.getRange("A2:D2").merge();
  overrideSheet.getRange("A2").setValue("Use this to permanently override categories for specific items in specific transactions. Applied BEFORE analysis.");
  overrideSheet.getRange("A2").setFontSize(10).setFontColor("#666666");
  
  overrideSheet.getRange("A3:D3").merge();
  overrideSheet.getRange("A3").setValue("Add rows with Transaction ID and Item Name to override. Wildcards (*) supported for Item Name.");
  overrideSheet.getRange("A3").setFontSize(9).setFontStyle("italic").setBackground("#f3e5f5");
  
  var headers = ["Transaction ID", "Item Name (or * for all items)", "New Category", "Notes"];
  overrideSheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  overrideSheet.getRange(5, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8E8E8");
  
  overrideSheet.getRange("A6").setValue("Valid Categories:");
  overrideSheet.getRange("A6").setFontWeight("bold");
  
  var validCategories = [
    ["Food"],
    ["Beverage"],
    ["Golf"],
    ["Membership"],
    ["Event"],
    ["Other"],
    ["Bay Rental"],
    ["Pizzas"],
    ["Or ANY category name from Square"]
  ];
  
  overrideSheet.getRange(7, 1, validCategories.length, 1).setValues(validCategories);
  overrideSheet.getRange(7, 1, validCategories.length, 1).setBackground("#f3e5f5");
  
  overrideSheet.getRange("A16").setValue("HOW TO FIND TRANSACTION ID:");
  overrideSheet.getRange("A16").setFontWeight("bold").setBackground("#fff9c4");
  
  var instructions = [
    ["1. Go to 'Square Item Detail Export' sheet"],
    ["2. Find the transaction you want to override"],
    ["3. Copy the Transaction ID from that row"],
    ["4. Copy the Item name (or use * for all items in that transaction)"],
    ["5. Paste here with the correct category"]
  ];
  
  overrideSheet.getRange(17, 1, instructions.length, 1).setValues(instructions);
  overrideSheet.getRange(17, 1, instructions.length, 1).setFontStyle("italic");
  
  overrideSheet.getRange("A23").setValue("EXAMPLE ENTRIES:");
  overrideSheet.getRange("A23").setFontWeight("bold").setBackground("#e1f5fe");
  
  var examples = [
    ["ABC123", "Bay Rental", "Golf", "This transaction was miscategorized"],
    ["XYZ789", "*", "Event", "Entire transaction should be Event"],
    ["DEF456", "Pepsi", "Beverage", "Individual item fix"],
    ["GHI789", "Monthly Membership", "Membership", "Membership fee"]
  ];
  
  overrideSheet.getRange(24, 1, examples.length, 4).setValues(examples);
  overrideSheet.getRange(24, 1, examples.length, 4).setBackground("#f5f5f5").setFontStyle("italic");
  
  overrideSheet.setColumnWidth(1, 200);
  overrideSheet.setColumnWidth(2, 250);
  overrideSheet.setColumnWidth(3, 150);
  overrideSheet.setColumnWidth(4, 300);
  overrideSheet.setFrozenRows(5);
  
  ss.setActiveSheet(overrideSheet);
  
  ui.alert("✅ Item Transaction Override Sheet Ready!",
    "Add rows below row 23 with:\n\n" +
    "Column A: Transaction ID (from Square Item Detail Export)\n" +
    "Column B: Item Name (or * for all items in transaction)\n" +
    "Column C: New Category\n" +
    "Column D: Optional notes\n\n" +
    "These overrides will be applied BEFORE every analysis runs.\n\n" +
    "Use this when specific transactions are always wrong in your imports.",
    ui.ButtonSet.OK);
}

function viewItemTransactionOverrides() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var overrideSheet = ss.getSheetByName("Item Transaction Override");
  
  if (!overrideSheet) {
    var response = ui.alert("No Override Sheet Found",
      "Would you like to create the Item Transaction Override sheet?",
      ui.ButtonSet.YES_NO);
    
    if (response === ui.Button.YES) {
      setupItemTransactionOverride();
    }
    return;
  }
  
  var overrideData = overrideSheet.getDataRange().getValues();
  var count = 0;
  var report = "ACTIVE ITEM TRANSACTION OVERRIDES:\n\n";
  
  for (var i = 1; i < overrideData.length; i++) {
    var transId = overrideData[i][0];
    var itemName = overrideData[i][1];
    var category = overrideData[i][2];
    var notes = overrideData[i][3];
    
    if (transId && category && !String(transId).includes("EXAMPLE") && !String(transId).includes("ABC123")) {
      count++;
      report += count + ". Transaction " + transId;
      if (itemName === "*") {
        report += " (ALL items)";
      } else {
        report += " / \"" + itemName + "\"";
      }
      report += " → " + category;
      if (notes) {
        report += "\n   (" + notes + ")";
      }
      report += "\n";
    }
  }
  
  if (count === 0) {
    report += "No active overrides found.\n\n";
    report += "Add overrides to the 'Item Transaction Override' sheet to permanently fix individual transactions.";
  } else {
    report += "\nTotal overrides: " + count;
  }
  
  ui.alert("Item Transaction Overrides", report, ui.ButtonSet.OK);
}

function applyItemTransactionOverrides() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var overrideSheet = ss.getSheetByName("Item Transaction Override");
  if (!overrideSheet) {
    return 0;
  }
  
  var overrideData = overrideSheet.getDataRange().getValues();
  
  var overrides = [];
  for (var i = 1; i < overrideData.length; i++) {
    var transId = overrideData[i][0];
    var itemName = overrideData[i][1];
    var category = overrideData[i][2];
    
    if (transId && category && !String(transId).includes("EXAMPLE") && !String(transId).includes("ABC123")) {
      overrides.push({
        transId: String(transId).trim(),
        itemName: itemName ? String(itemName).trim() : "*",
        category: String(category).trim()
      });
    }
  }
  
  if (overrides.length === 0) {
    return 0;
  }
  
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  if (!itemSheet) {
    return 0;
  }
  
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var transIdCol = itemHeaders.indexOf("Transaction ID");
  var itemNameCol = itemHeaders.indexOf("Item");
  var categoryCol = itemHeaders.indexOf("Category");
  
  if (transIdCol === -1 || categoryCol === -1) {
    return 0;
  }
  
  var changesApplied = 0;
  
  for (var i = 1; i < itemData.length; i++) {
    var transId = String(itemData[i][transIdCol]).trim();
    var itemName = itemNameCol !== -1 ? String(itemData[i][itemNameCol]).trim() : "";
    
    for (var j = 0; j < overrides.length; j++) {
      var override = overrides[j];
      
      var transMatches = (override.transId === transId);
      var itemMatches = (override.itemName === "*" || 
                        override.itemName.toLowerCase() === itemName.toLowerCase());
      
      if (transMatches && itemMatches) {
        itemSheet.getRange(i + 1, categoryCol + 1).setValue(override.category);
        changesApplied++;
        break;
      }
    }
  }
  
  return changesApplied;
}

// ============================================
// BOOKING PAYMENT VALIDATION
// ============================================

function validateBookingsVsPayments() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var validationSheet = ss.getSheetByName("Booking Payment Validation");
  if (!validationSheet) {
    validationSheet = ss.insertSheet("Booking Payment Validation");
  }
  
  validationSheet.clear();
  
  validationSheet.getRange("A1:L1").merge();
  validationSheet.getRange("A1").setValue("🚨 BOOKING vs PAYMENT VALIDATION");
  validationSheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  validationSheet.getRange("A1").setBackground("#1a1a1a").setFontColor("white");
  
  validationSheet.getRange("A2:L2").merge();
  validationSheet.getRange("A2").setValue("Flags bookings where payment doesn't match expected amount based on time/duration");
  validationSheet.getRange("A2").setFontSize(10).setHorizontalAlignment("center").setBackground("#f3f3f3");
  
  validationSheet.getRange("A3:L3").merge();
  validationSheet.getRange("A3").setValue("RATES: Weekday before 5pm = $25/hour | Weekday after 6pm & Weekends = $50/hour | 5pm-6pm transition uses both rates");
  validationSheet.getRange("A3").setFontSize(9).setHorizontalAlignment("center").setBackground("#fff3cd");
  
  var headers = [
    "Status",
    "Customer Name",
    "Email",
    "Booking Date",
    "Booking Time",
    "Duration (mins)",
    "Expected Amount",
    "Actual Paid",
    "Difference",
    "Transaction ID",
    "Booking ID",
    "Notes"
  ];
  
  validationSheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  validationSheet.getRange(5, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8E8E8");
  validationSheet.setFrozenRows(5);
  
  var bookingSheet = ss.getSheetByName("Apex Bookings Export");
  if (!bookingSheet) {
    ui.alert("Error", "Cannot find 'Apex Bookings Export' sheet!", ui.ButtonSet.OK);
    return;
  }
  
  var bookingData = bookingSheet.getDataRange().getValues();
  var bookingHeaders = bookingData[0];
  
  var bookingEmailCol = bookingHeaders.indexOf("Email");
  var bookingDateCol = bookingHeaders.indexOf("Date");
  var bookingTimeCol = bookingHeaders.indexOf("Time");
  var bookingDurationCol = bookingHeaders.indexOf("Duration Mins");
  var bookingIdCol = bookingHeaders.indexOf("Booking ID");
  var bookingFirstCol = bookingHeaders.indexOf("First Name");
  var bookingLastCol = bookingHeaders.indexOf("Last Name");
  
  var transSheet = ss.getSheetByName("Square Transactions Export");
  if (!transSheet) {
    ui.alert("Error", "Cannot find 'Square Transactions Export' sheet!", ui.ButtonSet.OK);
    return;
  }
  
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  
  var transIdCol = transHeaders.indexOf("Transaction ID");
  var transEmailCol = transHeaders.indexOf("Customer Email");
  var transNameCol = transHeaders.indexOf("Customer Name");
  var transDateCol = transHeaders.indexOf("Date");
  var transCollectedCol = transHeaders.indexOf("Total Collected");
  
  // Build transaction lookup by multiple keys for better matching
  var transactionsByEmail = {};
  var transactionsByName = {};
  var transactionsByDate = {};
  
  for (var i = 1; i < transData.length; i++) {
    var email = normalizeEmail(transData[i][transEmailCol]);
    var name = String(transData[i][transNameCol] || "").toLowerCase().trim();
    var date = transData[i][transDateCol];
    var amount = parseFloat(transData[i][transCollectedCol]) || 0;
    var transId = transData[i][transIdCol];
    
    if (!date) continue;
    
    var transObj = {
      amount: amount,
      transId: transId,
      date: date,
      email: email,
      name: name
    };
    
    // Index by email
    if (email) {
      if (!transactionsByEmail[email]) {
        transactionsByEmail[email] = [];
      }
      transactionsByEmail[email].push(transObj);
    }
    
    // Index by name
    if (name) {
      if (!transactionsByName[name]) {
        transactionsByName[name] = [];
      }
      transactionsByName[name].push(transObj);
    }
    
    // Index by date (for fallback)
    var dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
    if (!transactionsByDate[dateStr]) {
      transactionsByDate[dateStr] = [];
    }
    transactionsByDate[dateStr].push(transObj);
  }
  
  var results = [];
  var exactMatches = 0;
  var closeMatches = 0;
  var overcharged = 0;
  var undercharged = 0;
  var noPayment = 0;
  
  for (var i = 1; i < bookingData.length; i++) {
    var email = normalizeEmail(bookingData[i][bookingEmailCol]);
    var bookingDate = bookingData[i][bookingDateCol];
    var bookingTime = bookingData[i][bookingTimeCol];
    var duration = parseFloat(bookingData[i][bookingDurationCol]) || 0;
    var bookingId = bookingData[i][bookingIdCol];
    var firstName = bookingData[i][bookingFirstCol] || "";
    var lastName = bookingData[i][bookingLastCol] || "";
    var customerName = (firstName + " " + lastName).trim();
    
    if (!email || !bookingDate || !bookingTime || duration === 0) {
      continue;
    }
    
    var expectedAmount = calculateExpectedAmount(bookingDate, bookingTime, duration);
    
    // Try to find matching transaction(s)
    var candidateTransactions = [];
    
    // Method 1: Match by email (best)
    if (email && transactionsByEmail[email]) {
      candidateTransactions = candidateTransactions.concat(transactionsByEmail[email]);
    }
    
    // Method 2: Match by name (good fallback)
    var bookingNameKey = (firstName + " " + lastName).toLowerCase().trim();
    if (bookingNameKey && transactionsByName[bookingNameKey]) {
      candidateTransactions = candidateTransactions.concat(transactionsByName[bookingNameKey]);
    }
    
    // Method 3: Search nearby dates (±3 days) if still no match
    if (candidateTransactions.length === 0) {
      for (var dayOffset = -3; dayOffset <= 3; dayOffset++) {
        var searchDate = new Date(bookingDate);
        searchDate.setDate(searchDate.getDate() + dayOffset);
        var searchDateStr = Utilities.formatDate(searchDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        
        if (transactionsByDate[searchDateStr]) {
          candidateTransactions = candidateTransactions.concat(transactionsByDate[searchDateStr]);
        }
      }
    }
    
    // Filter candidates to within ±3 days of booking
    var relevantTransactions = [];
    var bookingTime = new Date(bookingDate).getTime();
    var threeDaysMs = 3 * 24 * 60 * 60 * 1000;
    
    for (var j = 0; j < candidateTransactions.length; j++) {
      var trans = candidateTransactions[j];
      var transTime = new Date(trans.date).getTime();
      
      if (Math.abs(transTime - bookingTime) <= threeDaysMs) {
        relevantTransactions.push(trans);
      }
    }
    
    var actualPaid = 0;
    var transId = "";
    var status = "";
    var difference = 0;
    var notes = "";
    
    if (relevantTransactions.length > 0) {
      // Find best match by amount closest to expected
      var bestMatch = null;
      var smallestDiff = 999999;
      
      for (var j = 0; j < relevantTransactions.length; j++) {
        var diff = Math.abs(relevantTransactions[j].amount - expectedAmount);
        if (diff < smallestDiff) {
          smallestDiff = diff;
          bestMatch = relevantTransactions[j];
        }
      }
      
      if (bestMatch) {
        actualPaid = bestMatch.amount;
        transId = bestMatch.transId;
        difference = actualPaid - expectedAmount;
        
        // Check date difference
        var daysDiff = Math.round((new Date(bestMatch.date) - new Date(bookingDate)) / (1000 * 60 * 60 * 24));
        var dateNote = daysDiff === 0 ? "" : " (paid " + (daysDiff > 0 ? daysDiff + " days after" : Math.abs(daysDiff) + " days before") + ")";
        
        if (difference === 0) {
          status = "✅ Exact Match";
          notes = dateNote;
          exactMatches++;
        } else if (difference > 0) {
          status = "💚 Overcharged";
          notes = "Paid $" + Math.abs(difference).toFixed(2) + " more than expected" + dateNote;
          overcharged++;
        } else if (difference >= -25) {
          status = "⚠️ Close Match";
          notes = "Undercharged by $" + Math.abs(difference).toFixed(2) + dateNote;
          closeMatches++;
        } else {
          status = "🚨 Underpaid";
          notes = "Undercharged by $" + Math.abs(difference).toFixed(2) + dateNote;
          undercharged++;
        }
        
        if (relevantTransactions.length > 1) {
          notes += " (" + relevantTransactions.length + " possible matches)";
        }
      }
    } else {
      status = "❌ No Payment";
      notes = "No matching transaction found within ±3 days";
      difference = -expectedAmount;
      noPayment++;
    }
    
    var timeStr = typeof bookingTime === 'string' ? bookingTime : 
                  Utilities.formatDate(bookingTime, Session.getScriptTimeZone(), "HH:mm");
    
    results.push([
      status,
      customerName || email,
      email,
      bookingDate,
      timeStr,
      duration,
      expectedAmount,
      actualPaid || "",
      difference !== 0 ? difference : "",
      transId,
      bookingId,
      notes
    ]);
  }
  
  results.sort(function(a, b) {
    var statusOrder = {"❌ No Payment": 1, "🚨 Underpaid": 2, "⚠️ Close Match": 3, "💚 Overcharged": 4, "✅ Exact Match": 5};
    return (statusOrder[a[0]] || 99) - (statusOrder[b[0]] || 99);
  });
  
  if (results.length > 0) {
    validationSheet.getRange(6, 1, results.length, headers.length).setValues(results);
    
    for (var i = 0; i < results.length; i++) {
      var row = i + 6;
      var status = results[i][0];
      var color;
      
      if (status === "❌ No Payment") {
        color = "#f4cccc";
      } else if (status === "🚨 Underpaid") {
        color = "#ea9999";
      } else if (status === "⚠️ Close Match") {
        color = "#fff2cc";
      } else if (status === "💚 Overcharged") {
        color = "#d9ead3";
      } else {
        color = "#ffffff";
      }
      
      validationSheet.getRange(row, 1, 1, headers.length).setBackground(color);
    }
    
    validationSheet.getRange(6, 7, results.length, 1).setNumberFormat("$#,##0.00");
    validationSheet.getRange(6, 8, results.length, 1).setNumberFormat("$#,##0.00");
    validationSheet.getRange(6, 9, results.length, 1).setNumberFormat("$#,##0.00");
    
    validationSheet.getRange(5, 1, results.length + 1, headers.length).setBorder(true, true, true, true, true, true);
  }
  
  validationSheet.getRange("A4").setValue("SUMMARY:");
  validationSheet.getRange("B4").setValue("✅ Exact: " + exactMatches);
  validationSheet.getRange("C4").setValue("💚 Over: " + overcharged);
  validationSheet.getRange("D4").setValue("⚠️ Close: " + closeMatches);
  validationSheet.getRange("E4").setValue("🚨 Under: " + undercharged);
  validationSheet.getRange("F4").setValue("❌ No Payment: " + noPayment);
  validationSheet.getRange("A4:F4").setFontWeight("bold").setBackground("#e8f0fe");
  
  for (var i = 1; i <= headers.length; i++) {
    validationSheet.autoResizeColumn(i);
  }
  
  ss.setActiveSheet(validationSheet);
  
  var summary = "Booking Payment Validation Complete!\n\n" +
    "✅ Exact Matches: " + exactMatches + "\n" +
    "💚 Overcharged: " + overcharged + "\n" +
    "⚠️ Close Matches (within $25 under): " + closeMatches + "\n" +
    "🚨 Underpaid (>$25 under): " + undercharged + "\n" +
    "❌ No Payment Found: " + noPayment + "\n\n" +
    "Total Bookings Checked: " + results.length;
  
  ui.alert("✅ Validation Complete", summary, ui.ButtonSet.OK);
}

function calculateExpectedAmount(date, time, durationMins) {
  var OFF_PEAK_RATE = 25;
  var PEAK_RATE = 50;
  
  var hour, minute;
  if (typeof time === 'string') {
    var parts = time.split(':');
    hour = parseInt(parts[0]);
    minute = parseInt(parts[1]) || 0;
  } else {
    hour = time.getHours();
    minute = time.getMinutes();
  }
  
  var dayOfWeek = new Date(date).getDay();
  var isWeekend = (dayOfWeek === 0 || dayOfWeek === 6);
  
  var totalCost = 0;
  var currentHour = hour;
  var currentMinute = minute;
  var remainingMins = durationMins;
  
  while (remainingMins > 0) {
    var currentRate;
    
    if (isWeekend) {
      currentRate = PEAK_RATE;
    } else {
      if (currentHour < 17) {
        currentRate = OFF_PEAK_RATE;
      } else if (currentHour >= 18) {
        currentRate = PEAK_RATE;
      } else {
        if (currentMinute === 0) {
          currentRate = OFF_PEAK_RATE;
        } else {
          currentRate = PEAK_RATE;
        }
      }
    }
    
    var minsUntilNextHour = 60 - currentMinute;
    var minsAtThisRate = Math.min(remainingMins, minsUntilNextHour);
    
    totalCost += (minsAtThisRate / 60) * currentRate;
    
    remainingMins -= minsAtThisRate;
    currentMinute = (currentMinute + minsAtThisRate) % 60;
    if (currentMinute === 0 && remainingMins > 0) {
      currentHour++;
    }
  }
  
  return Math.round(totalCost * 100) / 100;
}