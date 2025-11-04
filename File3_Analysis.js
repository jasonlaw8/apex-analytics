/**
 * ========================================
 * APEX GOLF ANALYTICS - FILE 3 OF 4
 * ========================================
 * 
 * THIS FILE CONTAINS:
 * - Top spenders functions (with date filtering)
 * - Individual analysis runners
 * - Master data population
 * - Data cleanup and categorization functions
 * 
 * REQUIRES: File 1 for date range and isDateInRange()
 * REQUIRES: File 2 for dashboard functions
 * REQUIRES: File 4 for helper functions
 */

// ============================================
// TOP SPENDERS FUNCTIONS
// ============================================

/**
 * Get top spenders excluding events (with date filtering)
 */
function getTopSpenders() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  
  var customerIdCol = transHeaders.indexOf("Customer ID");
  var collectedCol = transHeaders.indexOf("Total Collected");
  var transIdCol = transHeaders.indexOf("Transaction ID");
  var transDateCol = transHeaders.indexOf("Date");
  
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  
  var eventTransactions = {};
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    var category = itemData[i][itemCategoryCol];
    if (category && String(category).toLowerCase().trim() === "event") {
      eventTransactions[transId] = true;
    }
  }
  
  var customerSpending = {};
  
  for (var i = 1; i < transData.length; i++) {
    var date = transData[i][transDateCol];
    
    if (!isDateInRange(date)) {
      continue;
    }
    
    var customerId = transData[i][customerIdCol];
    var revenue = parseFloat(transData[i][collectedCol]) || 0;
    var transId = transData[i][transIdCol];
    
    if (customerId && !eventTransactions[transId]) {
      customerSpending[customerId] = (customerSpending[customerId] || 0) + revenue;
    }
  }
  
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  var custIdCol = customerHeaders.indexOf("Square Customer ID");
  var custFirstCol = customerHeaders.indexOf("First Name");
  var custLastCol = customerHeaders.indexOf("Last Name");
  var custEmailCol = customerHeaders.indexOf("Email Address");
  
  var customerIdToName = {};
  for (var i = 1; i < customerData.length; i++) {
    var id = customerData[i][custIdCol];
    var first = customerData[i][custFirstCol] || "";
    var last = customerData[i][custLastCol] || "";
    var email = customerData[i][custEmailCol] || "";
    var name = (first + " " + last).trim() || email || "Unknown";
    customerIdToName[id] = name;
  }
  
  var topSpenders = [];
  for (var customerId in customerSpending) {
    var name = customerIdToName[customerId] || customerId;
    topSpenders.push({name: name, spend: customerSpending[customerId].toFixed(2)});
  }
  topSpenders.sort(function(a, b) { return parseFloat(b.spend) - parseFloat(a.spend); });
  
  return topSpenders.slice(0, 10);
}

/**
 * Get top spenders excluding events AND memberships
 */
function getTopSpendersExcludingMembershipsEvents() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  
  if (!transSheet || !itemSheet) {
    return [];
  }
  
  var transData = transSheet.getDataRange().getValues();
  var itemData = itemSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  var itemHeaders = itemData[0];
  
  var transCustomerIdCol = transHeaders.indexOf("Customer ID");
  var transIdCol = transHeaders.indexOf("Transaction ID");
  var transDateCol = transHeaders.indexOf("Date");
  
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  var itemGrossCol = itemHeaders.indexOf("Gross Sales");
  var itemNameCol = itemHeaders.indexOf("Item");
  
  // Get overrides ONCE at the start - they're now cached internally
  var categoryOverrides = getCategoryOverrides();
  var transOverrides = getTransactionOverrides();
  
  // Build transaction to customer map and date map
  var transactionToCustomer = {};
  var transactionDates = {};
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transIdCol];
    var customerId = transData[i][transCustomerIdCol];
    var date = transData[i][transDateCol];
    
    if (transId && customerId) {
      transactionToCustomer[transId] = customerId;
      transactionDates[transId] = date;
    }
  }
  
  // Calculate spending per customer (excluding Events and Memberships)
  var customerSpend = {};
  
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    var itemName = itemData[i][itemNameCol];
    var category = itemData[i][itemCategoryCol];
    var grossSales = parseFloat(itemData[i][itemGrossCol]) || 0;
    
    // Check date range (uses cached settings)
    if (!isDateInRange(transactionDates[transId])) {
      continue;
    }
    
    // Apply overrides (from cached maps)
    if (transOverrides[transId]) {
      category = transOverrides[transId];
    } else if (itemName && categoryOverrides[itemName.toLowerCase()]) {
      category = categoryOverrides[itemName.toLowerCase()];
    }
    
    // getMajorCategory now checks item name first, then category
    var majorCat = getMajorCategory(category, itemName);
    
    // Skip Events and Memberships
    if (majorCat === "Event" || majorCat === "Membership") {
      continue;
    }
    
    // Get customer ID for this transaction
    var customerId = transactionToCustomer[transId];
    if (!customerId) {
      continue;
    }
    
    if (!customerSpend[customerId]) {
      customerSpend[customerId] = 0;
    }
    customerSpend[customerId] += grossSales;
  }
  
  // Get customer names
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  var custIdCol = customerHeaders.indexOf("Square Customer ID");
  var custFirstCol = customerHeaders.indexOf("First Name");
  var custLastCol = customerHeaders.indexOf("Last Name");
  var custEmailCol = customerHeaders.indexOf("Email Address");
  
  var customerIdToName = {};
  for (var i = 1; i < customerData.length; i++) {
    var id = customerData[i][custIdCol];
    var first = customerData[i][custFirstCol] || "";
    var last = customerData[i][custLastCol] || "";
    var email = customerData[i][custEmailCol] || "";
    var name = (first + " " + last).trim() || email || "Unknown";
    customerIdToName[id] = name;
  }
  
  // Build final list
  var customers = [];
  for (var customerId in customerSpend) {
    customers.push({
      name: customerIdToName[customerId] || customerId,
      spend: customerSpend[customerId].toFixed(2)
    });
  }
  
  customers.sort(function(a, b) {
    return parseFloat(b.spend) - parseFloat(a.spend);
  });
  
  return customers.slice(0, 10);
}

/**
 * Get top spenders by specific category
 */
function getTopSpendersByCategory(category) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  
  if (!transSheet || !itemSheet) {
    return [];
  }
  
  var transData = transSheet.getDataRange().getValues();
  var itemData = itemSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  var itemHeaders = itemData[0];
  
  var transCustomerIdCol = transHeaders.indexOf("Customer ID");
  var transIdCol = transHeaders.indexOf("Transaction ID");
  var transDateCol = transHeaders.indexOf("Date");
  
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  var itemGrossCol = itemHeaders.indexOf("Gross Sales");
  var itemNameCol = itemHeaders.indexOf("Item");
  
  var categoryOverrides = getCategoryOverrides();
  var transOverrides = getTransactionOverrides();
  
  // Build transaction to customer map and date map
  var transactionToCustomer = {};
  var transactionDates = {};
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transIdCol];
    var customerId = transData[i][transCustomerIdCol];
    var date = transData[i][transDateCol];
    
    if (transId && customerId) {
      transactionToCustomer[transId] = customerId;
      transactionDates[transId] = date;
    }
  }
  
  // Calculate spending per customer for this category
  var customerSpend = {};
  
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    var itemName = itemData[i][itemNameCol];
    var itemCategory = itemData[i][itemCategoryCol];
    var grossSales = parseFloat(itemData[i][itemGrossCol]) || 0;
    
    // Check date range
    if (!isDateInRange(transactionDates[transId])) {
      continue;
    }
    
    // Apply overrides
    if (transOverrides[transId]) {
      itemCategory = transOverrides[transId];
    } else if (itemName && categoryOverrides[itemName.toLowerCase()]) {
      itemCategory = categoryOverrides[itemName.toLowerCase()];
    }
    
    var majorCat = getMajorCategory(itemCategory, itemName);
    
    // For F&B category, check both Food and Beverage
    var matchesCategory = false;
    if (category === "F&B") {
      matchesCategory = (majorCat === "Food" || majorCat === "Beverage");
    } else {
      matchesCategory = (majorCat === category);
    }
    
    if (!matchesCategory) {
      continue;
    }
    
    // Get customer ID for this transaction
    var customerId = transactionToCustomer[transId];
    if (!customerId) {
      continue;
    }
    
    if (!customerSpend[customerId]) {
      customerSpend[customerId] = 0;
    }
    customerSpend[customerId] += grossSales;
  }
  
  // Get customer names
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  var custIdCol = customerHeaders.indexOf("Square Customer ID");
  var custFirstCol = customerHeaders.indexOf("First Name");
  var custLastCol = customerHeaders.indexOf("Last Name");
  var custEmailCol = customerHeaders.indexOf("Email Address");
  
  var customerIdToName = {};
  for (var i = 1; i < customerData.length; i++) {
    var id = customerData[i][custIdCol];
    var first = customerData[i][custFirstCol] || "";
    var last = customerData[i][custLastCol] || "";
    var email = customerData[i][custEmailCol] || "";
    var name = (first + " " + last).trim() || email || "Unknown";
    customerIdToName[id] = name;
  }
  
  // Build final list
  var customers = [];
  for (var customerId in customerSpend) {
    customers.push({
      name: customerIdToName[customerId] || customerId,
      spend: customerSpend[customerId].toFixed(2)
    });
  }
  
  customers.sort(function(a, b) {
    return parseFloat(b.spend) - parseFloat(a.spend);
  });
  
  return customers.slice(0, 10);
}

// ============================================
// INDIVIDUAL ANALYSIS RUNNERS
// ============================================

function runCustomerAnalysisOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Analytics Dashboard") || ss.insertSheet("Analytics Dashboard");
  sheet.clear();
  
  var dateSettings = getDateRangeSettings();
  
  sheet.getRange("A1").setValue("🏌️ CUSTOMER ANALYSIS");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Generated: " + new Date());
  
  var metrics = getCustomerMetrics();
  var currentRow = 4;
  
  var data = [
    ["Total Signups", metrics.totalSignups],
    ["Total Customers", metrics.totalCustomers],
    ["Signup to Customer %", metrics.signupToCustomerRate],
    ["Repeat Customers", metrics.repeatCustomers],
    ["One-Time Customers", metrics.oneTimeCustomers],
    ["Repeat Rate", metrics.repeatRate + "%"],
    ["Customer Booking Rate", metrics.customerBookingRate + "%"]
  ];
  
  sheet.getRange(currentRow, 1, data.length, 2).setValues(data);
  sheet.getRange(currentRow, 1, data.length, 2).setBorder(true, true, true, true, true, true);
  
  SpreadsheetApp.getUi().alert("✅ Complete!", "Check 'Analytics Dashboard'!", SpreadsheetApp.getUi().ButtonSet.OK);
}

function runRevenueAnalysisOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Analytics Dashboard") || ss.insertSheet("Analytics Dashboard");
  sheet.clear();
  
  var dateSettings = getDateRangeSettings();
  
  sheet.getRange("A1").setValue("🏌️ REVENUE ANALYSIS");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Generated: " + new Date());
  
  var metrics = getRevenueMetrics();
  var currentRow = 4;
  
  var data = [
    ["Total Revenue (w/ Events)", "$" + metrics.totalRevenue],
    ["Total Revenue (excl Events)", "$" + metrics.totalRevenueExclEvents],
    ["Event Revenue", "$" + metrics.eventRevenue],
    ["Total Net Revenue", "$" + metrics.totalNetRevenue],
    ["Total Tips", "$" + metrics.totalTips],
    ["Transactions (excl Events)", metrics.transactionsExclEvents],
    ["Avg Spend Per Visit", "$" + metrics.avgSpend],
    ["Avg Customer LTV", "$" + metrics.avgLTV]
  ];
  
  sheet.getRange(currentRow, 1, data.length, 2).setValues(data);
  sheet.getRange(currentRow, 1, data.length, 2).setBorder(true, true, true, true, true, true);
  
  SpreadsheetApp.getUi().alert("✅ Complete!", "Check 'Analytics Dashboard'!", SpreadsheetApp.getUi().ButtonSet.OK);
}

function runEnvisionAnalysisOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Analytics Dashboard") || ss.insertSheet("Analytics Dashboard");
  sheet.clear();
  
  var dateSettings = getDateRangeSettings();
  
  sheet.getRange("A1").setValue("🏌️ ENVISION RETENTION");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Generated: " + new Date());
  
  var metrics = getEnvisionMetrics();
  var currentRow = 4;
  
  var data = [
    ["Total Envision Customers", metrics.totalEnvision],
    ["Signups from Envision", metrics.signupsFromEnvision],
    ["Signup Retention Rate", metrics.signupRetentionRate + "%"],
    ["Customers from Envision", metrics.customersFromEnvision],
    ["Customer Retention Rate", metrics.customerRetentionRate + "%"],
    ["Revenue from Envision", "$" + metrics.revenueFromEnvision]
  ];
  
  sheet.getRange(currentRow, 1, data.length, 2).setValues(data);
  sheet.getRange(currentRow, 1, data.length, 2).setBorder(true, true, true, true, true, true);
  
  SpreadsheetApp.getUi().alert("✅ Complete!", "Check 'Analytics Dashboard'!", SpreadsheetApp.getUi().ButtonSet.OK);
}

function runBookingAnalysisOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Analytics Dashboard") || ss.insertSheet("Analytics Dashboard");
  sheet.clear();
  
  var dateSettings = getDateRangeSettings();
  
  sheet.getRange("A1").setValue("🏌️ BOOKING ANALYSIS");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Generated: " + new Date());
  
  var metrics = getBookingMetrics();
  var currentRow = 4;
  
  var data = [
    ["Total Bookings", metrics.totalBookings],
    ["Peak Hour", metrics.peakHour],
    ["Most Popular Day", metrics.popularDay],
    ["Avg Duration", metrics.avgDuration + " mins"]
  ];
  
  sheet.getRange(currentRow, 1, data.length, 2).setValues(data);
  sheet.getRange(currentRow, 1, data.length, 2).setBorder(true, true, true, true, true, true);
  
  SpreadsheetApp.getUi().alert("✅ Complete!", "Check 'Analytics Dashboard'!", SpreadsheetApp.getUi().ButtonSet.OK);
}

function runFoodBevAnalysisOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Analytics Dashboard") || ss.insertSheet("Analytics Dashboard");
  sheet.clear();
  
  var dateSettings = getDateRangeSettings();
  
  sheet.getRange("A1").setValue("🏌️ F&B ANALYSIS");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Generated: " + new Date());
  
  var metrics = getCategoryMetrics();
  var currentRow = 4;
  
  var data = [
    ["Food Revenue", "$" + metrics.foodRevenue],
    ["Beverage Revenue", "$" + metrics.beverageRevenue],
    ["Total F&B Revenue", "$" + metrics.totalFBRevenue],
    ["F&B % of Revenue", metrics.fbPercent + "%"],
    ["Avg F&B Per Transaction", "$" + metrics.avgFBPerTrans],
    ["F&B Attach Rate", metrics.fbAttachRate + "%"]
  ];
  
  sheet.getRange(currentRow, 1, data.length, 2).setValues(data);
  sheet.getRange(currentRow, 1, data.length, 2).setBorder(true, true, true, true, true, true);
  
  SpreadsheetApp.getUi().alert("✅ Complete!", "Check 'Analytics Dashboard'!", SpreadsheetApp.getUi().ButtonSet.OK);
}

function runBonusInsightsOnly() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName("Analytics Dashboard") || ss.insertSheet("Analytics Dashboard");
  sheet.clear();
  
  var dateSettings = getDateRangeSettings();
  
  sheet.getRange("A1").setValue("🏌️ BONUS INSIGHTS");
  sheet.getRange("A1").setFontSize(16).setFontWeight("bold");
  sheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Generated: " + new Date());
  
  SpreadsheetApp.getUi().alert("✅ Complete!", "Check 'Analytics Dashboard'!", SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================
// MASTER DATA POPULATION
// ============================================

function populateMasterData() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var masterSheet = ss.getSheetByName("Master Data");
  
  if (!masterSheet) {
    masterSheet = ss.insertSheet("Master Data");
  }
  
  masterSheet.clear();
  
  var dateSettings = getDateRangeSettings();
  
  masterSheet.getRange("A1").setValue("👥 MASTER DATA - Customer-Level Metrics");
  masterSheet.getRange("A1").setFontSize(14).setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  masterSheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Last Updated: " + new Date());
  
  var headers = [
    "Customer ID",
    "First Name",
    "Last Name",
    "Email",
    "Phone",
    "Date Added",
    "First Visit",
    "Last Visit",
    "# of Transactions",
    "Lifetime Spend",
    "# of Bookings",
    "Avg Days Between Visits",
    "Days Since First Visit",
    "Days Since Last Visit",
    "Previous Envision Customer?",
    "F&B Total Spend",
    "F&B % of Total Spend"
  ];
  
  masterSheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  masterSheet.getRange(4, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8E8E8");
  
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  
  var customerListSheet = ss.getSheetByName("Customer List");
  var customerListData = customerListSheet.getDataRange().getValues();
  var customerListHeaders = customerListData[0];
  
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  
  var bookingSheet = ss.getSheetByName("Apex Bookings Export");
  var bookingData = bookingSheet.getDataRange().getValues();
  var bookingHeaders = bookingData[0];
  
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var envisionSheet = ss.getSheetByName("Envision Customer List");
  var envisionData = envisionSheet.getDataRange().getValues();
  
  // Build Envision lookup
  var envisionByEmail = {};
  var envisionByPhone = {};
  var envisionByName = {};
  
  for (var i = 0; i < envisionData.length; i++) {
    var firstName = normalizeString(envisionData[i][0]);
    var lastName = normalizeString(envisionData[i][1]);
    var email = normalizeEmail(envisionData[i][2]);
    var phone = normalizePhone(envisionData[i][3]);
    
    if (email) envisionByEmail[email] = true;
    if (phone) envisionByPhone[phone] = true;
    if (firstName && lastName) envisionByName[firstName + "|" + lastName] = true;
  }
  
  var customerIdCol = customerHeaders.indexOf("Square Customer ID");
  var firstNameCol = customerHeaders.indexOf("First Name");
  var lastNameCol = customerHeaders.indexOf("Last Name");
  var emailCol = customerHeaders.indexOf("Email Address");
  var phoneCol = customerHeaders.indexOf("Phone Number");
  var firstVisitCol = customerHeaders.indexOf("First Visit");
  var lastVisitCol = customerHeaders.indexOf("Last Visit");
  var transCountCol = customerHeaders.indexOf("Transaction Count");
  var lifetimeSpendCol = customerHeaders.indexOf("Lifetime Spend");
  
  var clEmailCol = customerListHeaders.indexOf("Email");
  var clDateAddedCol = customerListHeaders.indexOf("Date Added");
  
  var customerListLookup = {};
  for (var i = 1; i < customerListData.length; i++) {
    var email = customerListData[i][clEmailCol];
    if (email) {
      customerListLookup[normalizeEmail(email)] = {
        dateAdded: customerListData[i][clDateAddedCol]
      };
    }
  }
  
  var bookingEmailCol = bookingHeaders.indexOf("Email");
  var bookingDateCol = bookingHeaders.indexOf("Date");
  var bookingsByEmail = {};
  for (var i = 1; i < bookingData.length; i++) {
    var email = bookingData[i][bookingEmailCol];
    var date = bookingData[i][bookingDateCol];
    
    if (!isDateInRange(date)) {
      continue;
    }
    
    if (email) {
      email = normalizeEmail(email);
      bookingsByEmail[email] = (bookingsByEmail[email] || 0) + 1;
    }
  }
  
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  var itemSalesCol = itemHeaders.indexOf("Gross Sales");
  var itemNameCol = itemHeaders.indexOf("Item");
  
  var transDateCol = transHeaders.indexOf("Date");
  var transTransIdCol = transHeaders.indexOf("Transaction ID");
  
  var transactionDates = {};
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transTransIdCol];
    var date = transData[i][transDateCol];
    transactionDates[transId] = date;
  }
  
  var fbByTransId = {};
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    
    if (!isDateInRange(transactionDates[transId])) {
      continue;
    }
    
    var category = itemData[i][itemCategoryCol];
    var sales = parseFloat(itemData[i][itemSalesCol]) || 0;
    var itemName = (itemData[i][itemNameCol] || "").toLowerCase();
    
    if (category === "Food" || category === "Beverage" ||
        itemName.includes("beer") || itemName.includes("wine") || itemName.includes("drink")) {
      fbByTransId[transId] = (fbByTransId[transId] || 0) + sales;
    }
  }
  
  var transCustomerIdCol = transHeaders.indexOf("Customer ID");
  var transIdCol = transHeaders.indexOf("Transaction ID");
  
  var fbByCustomerId = {};
  var visitsByCustomerId = {};
  
  for (var i = 1; i < transData.length; i++) {
    var date = transData[i][transDateCol];
    
    if (!isDateInRange(date)) {
      continue;
    }
    
    var customerId = transData[i][transCustomerIdCol];
    var transId = transData[i][transIdCol];
    
    if (customerId) {
      if (fbByTransId[transId]) {
        fbByCustomerId[customerId] = (fbByCustomerId[customerId] || 0) + fbByTransId[transId];
      }
      
      if (date) {
        if (!visitsByCustomerId[customerId]) {
          visitsByCustomerId[customerId] = [];
        }
        visitsByCustomerId[customerId].push(date);
      }
    }
  }
  
  var masterRows = [];
  var today = new Date();
  
  for (var i = 1; i < customerData.length; i++) {
    var customerId = customerData[i][customerIdCol];
    var firstVisit = customerData[i][firstVisitCol];
    
    if (!isDateInRange(firstVisit)) {
      continue;
    }
    
    var firstName = customerData[i][firstNameCol] || "";
    var lastName = customerData[i][lastNameCol] || "";
    var email = customerData[i][emailCol] || "";
    var phone = customerData[i][phoneCol] || "";
    var lastVisit = customerData[i][lastVisitCol];
    var transCount = customerData[i][transCountCol] || 0;
    var lifetimeSpend = customerData[i][lifetimeSpendCol] || 0;
    
    var dateAdded = "";
    var normEmail = normalizeEmail(email);
    if (normEmail && customerListLookup[normEmail]) {
      dateAdded = customerListLookup[normEmail].dateAdded || "";
    }
    
    var numBookings = 0;
    if (normEmail && bookingsByEmail[normEmail]) {
      numBookings = bookingsByEmail[normEmail];
    }
    
    var avgDaysBetween = "";
    if (visitsByCustomerId[customerId] && visitsByCustomerId[customerId].length > 1) {
      var visits = visitsByCustomerId[customerId].sort(function(a,b){return a-b;});
      var totalDays = 0;
      for (var v = 1; v < visits.length; v++) {
        totalDays += (visits[v] - visits[v-1]) / (1000*60*60*24);
      }
      avgDaysBetween = (totalDays / (visits.length - 1)).toFixed(1);
    }
    
    var daysSinceFirst = "";
    var daysSinceLast = "";
    if (firstVisit && firstVisit instanceof Date) {
      daysSinceFirst = Math.floor((today - firstVisit) / (1000*60*60*24));
    }
    if (lastVisit && lastVisit instanceof Date) {
      daysSinceLast = Math.floor((today - lastVisit) / (1000*60*60*24));
    }
    
    var isEnvision = "";
    var normPhone = normalizePhone(phone);
    var normFirst = normalizeString(firstName);
    var normLast = normalizeString(lastName);
    
    if ((normEmail && envisionByEmail[normEmail]) ||
        (normPhone && envisionByPhone[normPhone]) ||
        (normFirst && normLast && envisionByName[normFirst + "|" + normLast])) {
      isEnvision = "Yes";
    } else {
      isEnvision = "No";
    }
    
    var fbSpend = fbByCustomerId[customerId] || 0;
    var fbPercent = lifetimeSpend > 0 ? fbSpend / lifetimeSpend : 0;
    
    masterRows.push([
      customerId,
      firstName,
      lastName,
      email,
      phone,
      dateAdded,
      firstVisit,
      lastVisit,
      transCount,
      lifetimeSpend,
      numBookings,
      avgDaysBetween,
      daysSinceFirst,
      daysSinceLast,
      isEnvision,
      fbSpend,
      fbPercent
    ]);
  }
  
  if (masterRows.length > 0) {
    masterSheet.getRange(5, 1, masterRows.length, headers.length).setValues(masterRows);
    
    masterSheet.getRange(5, 10, masterRows.length, 1).setNumberFormat("$#,##0.00");
    masterSheet.getRange(5, 16, masterRows.length, 1).setNumberFormat("$#,##0.00");
    masterSheet.getRange(5, 17, masterRows.length, 1).setNumberFormat("0.0%");
  }
  
  masterSheet.autoResizeColumns(1, headers.length);
  masterSheet.setFrozenRows(4);
  masterSheet.setFrozenColumns(3);
  var lastRow = masterSheet.getLastRow();
  if (lastRow > 4) {
    masterSheet.getRange(4, 1, lastRow - 3, headers.length).setBorder(true, true, true, true, true, true);
  }
}

// ============================================
// DATA CLEANUP FUNCTIONS
// ============================================

function runDataCleanup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var cleanupSheet = ss.getSheetByName("Data Cleanup");
  if (!cleanupSheet) {
    cleanupSheet = ss.insertSheet("Data Cleanup");
    setupCleanupSheet(cleanupSheet);
  }
  
  var cleanupData = cleanupSheet.getDataRange().getValues();
  var savedCategories = {};
  
  for (var i = 1; i < cleanupData.length; i++) {
    var itemName = cleanupData[i][0];
    var category = cleanupData[i][1];
    if (itemName && category) {
      savedCategories[itemName.toLowerCase().trim()] = category;
    }
  }
  
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  if (!itemSheet) {
    ui.alert("Error", "Cannot find 'Square Item Detail Export' sheet!", ui.ButtonSet.OK);
    return;
  }
  
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var categoryCol = itemHeaders.indexOf("Category");
  var itemNameCol = itemHeaders.indexOf("Item");
  
  if (categoryCol === -1 || itemNameCol === -1) {
    ui.alert("Error", "Cannot find 'Category' or 'Item' columns!", ui.ButtonSet.OK);
    return;
  }
  
  var uncategorizedItems = {};
  var totalRows = itemData.length - 1;
  
  for (var i = 1; i < itemData.length; i++) {
    var category = itemData[i][categoryCol];
    var itemName = itemData[i][itemNameCol];
    
    if (!itemName) continue;
    
    var needsCategorization = false;
    
    if (!category || category === "" || category === "Uncategorized" ||
        String(category).toLowerCase().trim() === "uncategorized" ||
        String(category).toLowerCase().trim() === "none") {
      needsCategorization = true;
    }
    
    if (needsCategorization) {
      var normalizedName = itemName.toLowerCase().trim();
      
      if (savedCategories[normalizedName]) {
        itemSheet.getRange(i + 1, categoryCol + 1).setValue(savedCategories[normalizedName]);
      } else {
        if (!uncategorizedItems[normalizedName]) {
          uncategorizedItems[normalizedName] = {
            displayName: itemName,
            count: 0
          };
        }
        uncategorizedItems[normalizedName].count++;
      }
    }
  }
  
  var itemsToCategorize = [];
  for (var key in uncategorizedItems) {
    itemsToCategorize.push({
      normalizedName: key,
      displayName: uncategorizedItems[key].displayName,
      count: uncategorizedItems[key].count
    });
  }
  
  itemsToCategorize.sort(function(a, b) { return b.count - a.count; });
  
  if (itemsToCategorize.length === 0) {
    ui.alert("✅ All Clean!",
      "All items are categorized!\n\nTotal items checked: " + totalRows,
      ui.ButtonSet.OK);
    return;
  }
  
  var intro = ui.alert("🧹 Data Cleanup Needed",
    "Found " + itemsToCategorize.length + " unique items that need categorization.\n\n" +
    "You'll be asked to categorize each one.\n\nReady to start?",
    ui.ButtonSet.OK_CANCEL);
  
  if (intro !== ui.Button.OK) {
    return;
  }
  
  var newCategories = [];
  var categorized = 0;
  var skipped = 0;
  
  for (var i = 0; i < itemsToCategorize.length; i++) {
    var item = itemsToCategorize[i];
    
    var category = showCategorizationDialog(ui, item.displayName, item.count, i + 1, itemsToCategorize.length);
    
    if (category === null) {
      skipped++;
      break;
    }
    
    if (category === "SKIP") {
      skipped++;
      continue;
    }
    
    savedCategories[item.normalizedName] = category;
    newCategories.push([item.displayName, category, new Date(), item.count]);
    categorized++;
    
    for (var j = 1; j < itemData.length; j++) {
      var rowItemName = itemData[j][itemNameCol];
      if (rowItemName && rowItemName.toLowerCase().trim() === item.normalizedName) {
        itemSheet.getRange(j + 1, categoryCol + 1).setValue(category);
      }
    }
  }
  
  if (newCategories.length > 0) {
    var lastRow = cleanupSheet.getLastRow();
    cleanupSheet.getRange(lastRow + 1, 1, newCategories.length, 4).setValues(newCategories);
  }
  
  var summary = "✅ Cleanup Complete!\n\n" +
    "Categorized: " + categorized + " items\n" +
    "Skipped: " + skipped + " items\n\n" +
    "Run cleanup again to categorize remaining items.";
  
  ui.alert("Summary", summary, ui.ButtonSet.OK);
}

function showCategorizationDialog(ui, itemName, count, current, total) {
  var choice = ui.prompt(
    "Categorize Item (" + current + " of " + total + ")",
    "Item: " + itemName + "\nAppears " + count + " times\n\n" +
    "Choose category:\n" +
    "1 = Food\n" +
    "2 = Beverage\n" +
    "3 = Bay Rental\n" +
    "4 = Event\n" +
    "5 = Other\n" +
    "S = Skip\n\n" +
    "Enter your choice:",
    ui.ButtonSet.OK_CANCEL
  );
  
  if (choice.getSelectedButton() !== ui.Button.OK) {
    return null;
  }
  
  var input = choice.getResponseText().trim().toUpperCase();
  
  switch(input) {
    case "1": return "Food";
    case "2": return "Beverage";
    case "3": return "Bay Rental";
    case "4": return "Event";
    case "5": return "Other";
    case "S": return "SKIP";
    default:
      ui.alert("Invalid input. Please try again.");
      return showCategorizationDialog(ui, itemName, count, current, total);
  }
}

function setupCleanupSheet(sheet) {
  sheet.clear();
  
  sheet.getRange("A1").setValue("🧹 DATA CLEANUP - Item Categories");
  sheet.getRange("A1").setFontSize(14).setFontWeight("bold").setBackground("#4285F4").setFontColor("white");
  sheet.getRange("A1:D1").merge();
  
  sheet.getRange("A2").setValue("This sheet stores your item categorizations. Do not edit manually.");
  sheet.getRange("A2:D2").merge();
  
  var headers = ["Item Name", "Category", "Date Categorized", "# of Occurrences"];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(4, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8E8E8");
  
  sheet.setFrozenRows(4);
  sheet.autoResizeColumns(1, 4);
}

function viewUncategorizedItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var cleanupSheet = ss.getSheetByName("Data Cleanup");
  var savedCategories = {};
  
  if (cleanupSheet) {
    var cleanupData = cleanupSheet.getDataRange().getValues();
    for (var i = 1; i < cleanupData.length; i++) {
      var itemName = cleanupData[i][0];
      var category = cleanupData[i][1];
      if (itemName && category) {
        savedCategories[itemName.toLowerCase().trim()] = category;
      }
    }
  }
  
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  if (!itemSheet) {
    ui.alert("Error", "Cannot find 'Square Item Detail Export' sheet!", ui.ButtonSet.OK);
    return;
  }
  
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var categoryCol = itemHeaders.indexOf("Category");
  var itemNameCol = itemHeaders.indexOf("Item");
  
  var uncategorized = {};
  
  for (var i = 1; i < itemData.length; i++) {
    var category = itemData[i][categoryCol];
    var itemName = itemData[i][itemNameCol];
    
    if (!itemName) continue;
    
    if (!category || category === "" || category === "Uncategorized" ||
        String(category).toLowerCase().trim() === "uncategorized" ||
        String(category).toLowerCase().trim() === "none") {
      var normalizedName = itemName.toLowerCase().trim();
      
      if (savedCategories[normalizedName]) continue;
      
      if (!uncategorized[normalizedName]) {
        uncategorized[normalizedName] = {
          displayName: itemName,
          count: 0
        };
      }
      uncategorized[normalizedName].count++;
    }
  }
  
  var report = "UNCATEGORIZED ITEMS REPORT\n\n";
  var items = [];
  
  for (var key in uncategorized) {
    items.push(uncategorized[key]);
  }
  
  items.sort(function(a, b) { return b.count - a.count; });
  
  if (items.length === 0) {
    report += "✅ All items are categorized!";
  } else {
    report += "Found " + items.length + " unique items:\n\n";
    
    for (var i = 0; i < Math.min(20, items.length); i++) {
      report += (i + 1) + ". " + items[i].displayName + " (" + items[i].count + " occurrences)\n";
    }
    
    if (items.length > 20) {
      report += "\n... and " + (items.length - 20) + " more items.";
    }
    
    report += "\n\nRun 'Clean Data' to categorize these items.";
  }
  
  ui.alert("Uncategorized Items", report, ui.ButtonSet.OK);
}