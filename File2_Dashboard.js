/**
 * ========================================
 * APEX GOLF ANALYTICS - FILE 2 OF 4
 * ========================================
 * 
 * THIS FILE CONTAINS:
 * - Main dashboard builder (runCompleteAnalysis)
 * - Core metric extraction functions with date filtering
 * - Dashboard formatting functions
 * - Top spenders functions
 * 
 * REQUIRES: File 1 for date range globals and isDateInRange()
 * REQUIRES: File 3 for analysis functions
 * REQUIRES: File 4 for helper functions
 */

// ============================================
// MAIN FUNCTION - COMPLETE ANALYSIS
// ============================================

function runCompleteAnalysis() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  Logger.log("Starting analysis...");
  
  // Get current date range settings
  var dateSettings = getDateRangeSettings();
  var ANALYSIS_RANGE_LABEL = dateSettings.label;
  
  Logger.log("Date range: " + ANALYSIS_RANGE_LABEL);
  
  // Apply item transaction overrides FIRST
  var overridesApplied = applyItemTransactionOverrides();
  
  Logger.log("Overrides applied: " + overridesApplied);
  
  if (overridesApplied > 0) {
    SpreadsheetApp.getUi().alert("✅ Overrides Applied",
      overridesApplied + " item transaction override(s) applied to Square Item Detail Export.\n\nProceeding with analysis...",
      SpreadsheetApp.getUi().ButtonSet.OK);
  }
  
  Logger.log("Creating dashboard sheet...");
  
  var dashboardSheet = ss.getSheetByName("Analytics Dashboard");
  if (dashboardSheet) {
    dashboardSheet.clear();
  } else {
    dashboardSheet = ss.insertSheet("Analytics Dashboard");
  }
  
  Logger.log("Setting up dashboard layout...");
  
  // Set up dashboard header
  dashboardSheet.setColumnWidth(1, 250);
  dashboardSheet.setColumnWidth(2, 150);
  dashboardSheet.setColumnWidth(3, 50);
  dashboardSheet.setColumnWidth(4, 250);
  dashboardSheet.setColumnWidth(5, 150);
  dashboardSheet.setColumnWidth(6, 50);
  dashboardSheet.setColumnWidth(7, 250);
  dashboardSheet.setColumnWidth(8, 150);
  
  // Main header
  dashboardSheet.getRange("A1:H1").merge();
  dashboardSheet.getRange("A1").setValue("🏌️ APEX GOLF ANALYTICS DASHBOARD");
  dashboardSheet.getRange("A1").setFontSize(18).setFontWeight("bold").setHorizontalAlignment("center");
  dashboardSheet.getRange("A1").setBackground("#1a1a1a").setFontColor("white");
  
  // Date range and timestamp
  dashboardSheet.getRange("A2:H2").merge();
  dashboardSheet.getRange("A2").setValue("📅 Date Range: " + ANALYSIS_RANGE_LABEL + " | Generated: " + new Date());
  dashboardSheet.getRange("A2").setFontSize(10).setHorizontalAlignment("center");
  dashboardSheet.getRange("A2").setBackground("#fff3cd");
  
  Logger.log("Getting customer metrics...");
  
  // Get all metrics (they now respect date range via isDateInRange())
  var customerMetrics = getCustomerMetrics();
  
  Logger.log("Getting Envision metrics...");
  var envisionMetrics = getEnvisionMetrics();
  
  Logger.log("Getting revenue metrics...");
  var revenueMetrics = getRevenueMetrics();
  
  Logger.log("Getting category metrics...");
  var categoryMetrics = getCategoryMetrics();
  
  Logger.log("Getting booking metrics...");
  var bookingMetrics = getBookingMetrics();
  
  var currentRow = 4;
  
  // ROW 1: Key metrics cards
  currentRow = createMetricCard(dashboardSheet, currentRow, 1, "👥 TOTAL SIGNUPS", customerMetrics.totalSignups, "#4285F4");
  createMetricCard(dashboardSheet, currentRow, 4, "💰 TOTAL REVENUE", "$" + revenueMetrics.totalRevenue, "#34A853");
  createMetricCard(dashboardSheet, currentRow, 7, "📅 TOTAL BOOKINGS", bookingMetrics.totalBookings, "#FBBC04");
  
  currentRow += 3;
  
  // ROW 2: Secondary metrics
  createMetricCard(dashboardSheet, currentRow, 1, "🔄 REPEAT RATE", customerMetrics.repeatRate + "%", "#4285F4");
  createMetricCard(dashboardSheet, currentRow, 4, "💵 AVG SPEND/VISIT", "$" + revenueMetrics.avgSpend, "#34A853");
  createMetricCard(dashboardSheet, currentRow, 7, "⭐ ENVISION RETENTION", envisionMetrics.customerRetentionRate + "%", "#EA4335");
  
  currentRow += 4;
  
  // CUSTOMER ANALYSIS - Left side
  dashboardSheet.getRange(currentRow, 1, 1, 2).merge();
  dashboardSheet.getRange(currentRow, 1).setValue("👥 CUSTOMER BASE");
  dashboardSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#4285F4").setFontColor("white");
  currentRow++;
  
  var customerData = [
    ["Total Signups", customerMetrics.totalSignups],
    ["Total Customers", customerMetrics.totalCustomers],
    ["Signup → Customer Rate", customerMetrics.signupToCustomerRate + "%"],
    ["Repeat Customers", customerMetrics.repeatCustomers],
    ["One-Time Customers", customerMetrics.oneTimeCustomers],
    ["Customer Booking Rate", customerMetrics.customerBookingRate + "%"]
  ];
  
  dashboardSheet.getRange(currentRow, 1, customerData.length, 2).setValues(customerData);
  dashboardSheet.getRange(currentRow, 1, customerData.length, 2).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(currentRow, 1, customerData.length, 1).setBackground("#e8f0fe");
  
  var customerEndRow = currentRow + customerData.length;
  
  // ENVISION RETENTION - Right side
  var envisionStartRow = currentRow - 1;
  dashboardSheet.getRange(envisionStartRow, 4, 1, 2).merge();
  dashboardSheet.getRange(envisionStartRow, 4).setValue("🔄 ENVISION → APEX");
  dashboardSheet.getRange(envisionStartRow, 4).setFontWeight("bold").setFontSize(12).setBackground("#34A853").setFontColor("white");
  
  var envisionData = [
    ["Total Envision Customers", envisionMetrics.totalEnvision],
    ["Signups from Envision", envisionMetrics.signupsFromEnvision],
    ["Signup Retention Rate", envisionMetrics.signupRetentionRate + "%"],
    ["Customers from Envision", envisionMetrics.customersFromEnvision],
    ["Customer Retention Rate", envisionMetrics.customerRetentionRate + "%"],
    ["Revenue from Envision", "$" + envisionMetrics.revenueFromEnvision]
  ];
  
  dashboardSheet.getRange(envisionStartRow + 1, 4, envisionData.length, 2).setValues(envisionData);
  dashboardSheet.getRange(envisionStartRow + 1, 4, envisionData.length, 2).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(envisionStartRow + 1, 4, envisionData.length, 1).setBackground("#e6f4ea");
  
  currentRow = Math.max(customerEndRow, envisionStartRow + 1 + envisionData.length) + 2;
  
  // REVENUE BREAKDOWN
  dashboardSheet.getRange(currentRow, 1, 1, 5).merge();
  dashboardSheet.getRange(currentRow, 1).setValue("💰 REVENUE BREAKDOWN");
  dashboardSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#FBBC04").setFontColor("white");
  currentRow++;
  
  var revenueData = [
    ["Total Revenue (w/ Events)", "$" + revenueMetrics.totalRevenue, "", "Event Revenue", "$" + revenueMetrics.eventRevenue],
    ["Total Revenue (excl Events)", "$" + revenueMetrics.totalRevenueExclEvents, "", "Avg Spend Per Visit", "$" + revenueMetrics.avgSpend],
    ["Total Net Revenue", "$" + revenueMetrics.totalNetRevenue, "", "Avg Customer LTV", "$" + revenueMetrics.avgLTV],
    ["Total Tips", "$" + revenueMetrics.totalTips, "", "Total Transactions (excl Events)", revenueMetrics.transactionsExclEvents]
  ];
  
  dashboardSheet.getRange(currentRow, 1, revenueData.length, 5).setValues(revenueData);
  dashboardSheet.getRange(currentRow, 1, revenueData.length, 5).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(currentRow, 1, revenueData.length, 1).setBackground("#fef7e0");
  dashboardSheet.getRange(currentRow, 4, revenueData.length, 1).setBackground("#fef7e0");
  
  currentRow += revenueData.length + 2;
  
  // CATEGORY SPLIT
  dashboardSheet.getRange(currentRow, 1, 1, 8).merge();
  dashboardSheet.getRange(currentRow, 1).setValue("📊 REVENUE BY MAJOR CATEGORY (Excluding Events)");
  dashboardSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#9C27B0").setFontColor("white");
  currentRow++;
  
  var categoryData = [
    ["🍔 Food", "$" + categoryMetrics.foodRevenue, categoryMetrics.foodPercent + "%", ""],
    ["🍺 Beverage", "$" + categoryMetrics.beverageRevenue, categoryMetrics.beveragePercent + "%", ""],
    ["⛳ Golf", "$" + categoryMetrics.golfRevenue, categoryMetrics.golfPercent + "%", ""],
    ["👤 Membership", "$" + categoryMetrics.membershipRevenue, categoryMetrics.membershipPercent + "%", ""],
    ["📦 Miscellaneous", "$" + categoryMetrics.miscRevenue, categoryMetrics.miscPercent + "%", ""]
  ];
  
  dashboardSheet.getRange(currentRow, 1, categoryData.length, 4).setValues(categoryData);
  dashboardSheet.getRange(currentRow, 1, categoryData.length, 4).setBorder(true, true, true, true, true, true);
  
  for (var i = 0; i < categoryData.length; i++) {
    var percent = parseFloat(categoryData[i][2]);
    var color = getColorForPercent(percent);
    dashboardSheet.getRange(currentRow + i, 1, 1, 3).setBackground(color);
  }
  
  currentRow += categoryData.length + 2;
  
  // F&B METRICS - Left side
  dashboardSheet.getRange(currentRow, 1, 1, 2).merge();
  dashboardSheet.getRange(currentRow, 1).setValue("🍔 F&B PERFORMANCE");
  dashboardSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#9C27B0").setFontColor("white");
  currentRow++;
  
  var fbData = [
    ["Total F&B Revenue", "$" + categoryMetrics.totalFBRevenue],
    ["F&B % of Revenue", categoryMetrics.fbPercent + "%"],
    ["Avg F&B Per Transaction", "$" + categoryMetrics.avgFBPerTrans],
    ["F&B Attach Rate", categoryMetrics.fbAttachRate + "%"]
  ];
  
  dashboardSheet.getRange(currentRow, 1, fbData.length, 2).setValues(fbData);
  dashboardSheet.getRange(currentRow, 1, fbData.length, 2).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(currentRow, 1, fbData.length, 1).setBackground("#f3e5f5");
  
  var fbEndRow = currentRow + fbData.length;
  
  // BOOKING METRICS - Right side
  var bookingStartRow = currentRow - 1;
  dashboardSheet.getRange(bookingStartRow, 4, 1, 2).merge();
  dashboardSheet.getRange(bookingStartRow, 4).setValue("📅 BOOKING INSIGHTS");
  dashboardSheet.getRange(bookingStartRow, 4).setFontWeight("bold").setFontSize(12).setBackground("#EA4335").setFontColor("white");
  
  var bookingData = [
    ["Total Bookings", bookingMetrics.totalBookings],
    ["Peak Hour", bookingMetrics.peakHour],
    ["Most Popular Day", bookingMetrics.popularDay],
    ["Avg Duration", bookingMetrics.avgDuration + " mins"]
  ];
  
  dashboardSheet.getRange(bookingStartRow + 1, 4, bookingData.length, 2).setValues(bookingData);
  dashboardSheet.getRange(bookingStartRow + 1, 4, bookingData.length, 2).setBorder(true, true, true, true, true, true);
  dashboardSheet.getRange(bookingStartRow + 1, 4, bookingData.length, 1).setBackground("#fce8e6");
  
  currentRow = Math.max(fbEndRow, bookingStartRow + 1 + bookingData.length) + 2;
  
  Logger.log("Getting top spenders...");
  
  // TOP SPENDERS
  dashboardSheet.getRange(currentRow, 1, 1, 8).merge();
  dashboardSheet.getRange(currentRow, 1).setValue("🏆 TOP 10 CUSTOMERS BY CATEGORY");
  dashboardSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#FF6D00").setFontColor("white");
  currentRow++;
  
  var headerRow = currentRow;
  dashboardSheet.getRange(headerRow, 1, 1, 2).merge();
  dashboardSheet.getRange(headerRow, 1).setValue("Excl. Events Only");
  dashboardSheet.getRange(headerRow, 1).setFontWeight("bold").setBackground("#ffe0b2").setHorizontalAlignment("center");
  
  dashboardSheet.getRange(headerRow, 3).setValue("");
  
  dashboardSheet.getRange(headerRow, 4, 1, 2).merge();
  dashboardSheet.getRange(headerRow, 4).setValue("Excl. Events + Memberships");
  dashboardSheet.getRange(headerRow, 4).setFontWeight("bold").setBackground("#ffcc80").setHorizontalAlignment("center");
  
  dashboardSheet.getRange(headerRow, 6).setValue("");
  
  dashboardSheet.getRange(headerRow, 7, 1, 2).merge();
  dashboardSheet.getRange(headerRow, 7).setValue("F&B Only");
  dashboardSheet.getRange(headerRow, 7).setFontWeight("bold").setBackground("#fff3e0").setHorizontalAlignment("center");
  
  currentRow++;
  
  var topSpenders = getTopSpenders();
  var topSpendersNoMembership = getTopSpendersExcludingMembershipsEvents();
  var topSpendersFB = getTopSpendersByCategory("F&B");
  
  var maxRows = Math.max(topSpenders.length, topSpendersNoMembership.length, topSpendersFB.length);
  
  for (var i = 0; i < maxRows; i++) {
    if (i < topSpenders.length) {
      dashboardSheet.getRange(currentRow + i, 1).setValue((i + 1) + ". " + topSpenders[i].name);
      dashboardSheet.getRange(currentRow + i, 2).setValue("$" + topSpenders[i].spend);
    }
    
    if (i < topSpendersNoMembership.length) {
      dashboardSheet.getRange(currentRow + i, 4).setValue((i + 1) + ". " + topSpendersNoMembership[i].name);
      dashboardSheet.getRange(currentRow + i, 5).setValue("$" + topSpendersNoMembership[i].spend);
    }
    
    if (i < topSpendersFB.length) {
      dashboardSheet.getRange(currentRow + i, 7).setValue((i + 1) + ". " + topSpendersFB[i].name);
      dashboardSheet.getRange(currentRow + i, 8).setValue("$" + topSpendersFB[i].spend);
    }
  }
  
  if (maxRows > 0) {
    if (topSpenders.length > 0) {
      dashboardSheet.getRange(currentRow, 1, topSpenders.length, 2).setBorder(true, true, true, true, true, true);
      if (topSpenders.length >= 1) dashboardSheet.getRange(currentRow, 1, 1, 2).setBackground("#fff3e0");
      if (topSpenders.length >= 2) dashboardSheet.getRange(currentRow + 1, 1, 1, 2).setBackground("#ffe0b2");
      if (topSpenders.length >= 3) dashboardSheet.getRange(currentRow + 2, 1, 1, 2).setBackground("#ffcc80");
    }
    
    if (topSpendersNoMembership.length > 0) {
      dashboardSheet.getRange(currentRow, 4, topSpendersNoMembership.length, 2).setBorder(true, true, true, true, true, true);
      if (topSpendersNoMembership.length >= 1) dashboardSheet.getRange(currentRow, 4, 1, 2).setBackground("#fff3e0");
      if (topSpendersNoMembership.length >= 2) dashboardSheet.getRange(currentRow + 1, 4, 1, 2).setBackground("#ffe0b2");
      if (topSpendersNoMembership.length >= 3) dashboardSheet.getRange(currentRow + 2, 4, 1, 2).setBackground("#ffcc80");
    }
    
    if (topSpendersFB.length > 0) {
      dashboardSheet.getRange(currentRow, 7, topSpendersFB.length, 2).setBorder(true, true, true, true, true, true);
      if (topSpendersFB.length >= 1) dashboardSheet.getRange(currentRow, 7, 1, 2).setBackground("#fff3e0");
      if (topSpendersFB.length >= 2) dashboardSheet.getRange(currentRow + 1, 7, 1, 2).setBackground("#ffe0b2");
      if (topSpendersFB.length >= 3) dashboardSheet.getRange(currentRow + 2, 7, 1, 2).setBackground("#ffcc80");
    }
  }
  
  currentRow += maxRows + 2;
  
  Logger.log("Top spenders complete!");
  
  Logger.log("Populating Master Data...");
  populateMasterData();
  Logger.log("Master Data complete!");
  
  SpreadsheetApp.getUi().alert("✅ Analysis Complete!",
    "Date Range: " + ANALYSIS_RANGE_LABEL + "\n\n" +
    "Check 'Analytics Dashboard' and 'Master Data' sheets!",
    SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Creates a metric card (KPI box)
 */
function createMetricCard(sheet, row, col, title, value, color) {
  sheet.getRange(row, col, 1, 2).merge();
  sheet.getRange(row, col).setValue(title);
  sheet.getRange(row, col).setFontSize(10).setFontWeight("bold").setBackground(color).setFontColor("white");
  sheet.getRange(row, col).setHorizontalAlignment("center");
  
  sheet.getRange(row + 1, col, 1, 2).merge();
  sheet.getRange(row + 1, col).setValue(value);
  sheet.getRange(row + 1, col).setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  sheet.getRange(row + 1, col).setBorder(true, true, true, true, false, false);
  
  return row;
}

/**
 * Gets color based on percentage for visual bars
 */
function getColorForPercent(percent) {
  if (percent >= 30) return "#c8e6c9";
  if (percent >= 20) return "#fff9c4";
  if (percent >= 10) return "#ffccbc";
  return "#f5f5f5";
}

// ============================================
// CORE METRIC EXTRACTORS (WITH DATE FILTERING)
// ============================================

/**
 * Extract customer metrics for dashboard
 */
function getCustomerMetrics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  var transactionCountCol = customerHeaders.indexOf("Transaction Count");
  var firstVisitCol = customerHeaders.indexOf("First Visit");
  
  var customerListSheet = ss.getSheetByName("Customer List");
  var customerListData = customerListSheet.getDataRange().getValues();
  
  var totalSignups = customerListData.length - 1;
  
  // Filter customers by date range
  var filteredCustomers = 0;
  var repeatCustomers = 0;
  var oneTimeCustomers = 0;
  
  for (var i = 1; i < customerData.length; i++) {
    var firstVisit = customerData[i][firstVisitCol];
    
    if (!isDateInRange(firstVisit)) {
      continue;
    }
    
    filteredCustomers++;
    
    var transCount = customerData[i][transactionCountCol];
    if (transCount > 1) {
      repeatCustomers++;
    } else if (transCount == 1) {
      oneTimeCustomers++;
    }
  }
  
  var bookingSheet = ss.getSheetByName("Apex Bookings Export");
  var bookingData = bookingSheet.getDataRange().getValues();
  var bookingHeaders = bookingData[0];
  var bookingEmailCol = bookingHeaders.indexOf("Email");
  var bookingDateCol = bookingHeaders.indexOf("Date");
  
  var uniqueCustomersWithBookings = {};
  for (var i = 1; i < bookingData.length; i++) {
    var email = bookingData[i][bookingEmailCol];
    var bookingDate = bookingData[i][bookingDateCol];
    
    if (!isDateInRange(bookingDate)) {
      continue;
    }
    
    if (email) {
      uniqueCustomersWithBookings[normalizeEmail(email)] = true;
    }
  }
  
  var customersWhoBooked = Object.keys(uniqueCustomersWithBookings).length;
  
  return {
    totalSignups: totalSignups,
    totalCustomers: filteredCustomers,
    signupToCustomerRate: (filteredCustomers / totalSignups * 100).toFixed(1),
    repeatCustomers: repeatCustomers,
    oneTimeCustomers: oneTimeCustomers,
    repeatRate: filteredCustomers > 0 ? (repeatCustomers / filteredCustomers * 100).toFixed(1) : "0.0",
    customerBookingRate: filteredCustomers > 0 ? (customersWhoBooked / filteredCustomers * 100).toFixed(1) : "0.0"
  };
}

/**
 * Extract Envision metrics for dashboard
 */
function getEnvisionMetrics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var envisionSheet = ss.getSheetByName("Envision Customer List");
  var envisionData = envisionSheet.getDataRange().getValues();
  
  var envisionByEmail = {};
  var envisionByPhone = {};
  var envisionByName = {};
  
  for (var i = 0; i < envisionData.length; i++) {
    var firstName = normalizeString(envisionData[i][0]);
    var lastName = normalizeString(envisionData[i][1]);
    var email = normalizeEmail(envisionData[i][2]);
    var phone = normalizePhone(envisionData[i][3]);
    
    var customerKey = i + "_" + firstName + "_" + lastName;
    
    if (email) envisionByEmail[email] = customerKey;
    if (phone) envisionByPhone[phone] = customerKey;
    if (firstName && lastName) {
      envisionByName[firstName + "|" + lastName] = customerKey;
    }
  }
  
  var customerListSheet = ss.getSheetByName("Customer List");
  var customerListData = customerListSheet.getDataRange().getValues();
  var customerListHeaders = customerListData[0];
  var clEmailCol = customerListHeaders.indexOf("Email");
  var clPhoneCol = customerListHeaders.indexOf("Phone Number");
  var clFirstCol = customerListHeaders.indexOf("First Name");
  var clLastCol = customerListHeaders.indexOf("Last Name");
  
  var matchedEnvisionSignups = new Set();
  
  for (var i = 1; i < customerListData.length; i++) {
    var email = normalizeEmail(customerListData[i][clEmailCol]);
    var phone = normalizePhone(customerListData[i][clPhoneCol]);
    var firstName = normalizeString(customerListData[i][clFirstCol]);
    var lastName = normalizeString(customerListData[i][clLastCol]);
    
    var matchedKey = null;
    
    if (email && envisionByEmail[email]) {
      matchedKey = envisionByEmail[email];
    } else if (phone && envisionByPhone[phone]) {
      matchedKey = envisionByPhone[phone];
    } else if (firstName && lastName) {
      var nameKey = firstName + "|" + lastName;
      if (envisionByName[nameKey]) {
        matchedKey = envisionByName[nameKey];
      }
    }
    
    if (matchedKey) {
      matchedEnvisionSignups.add(matchedKey);
    }
  }
  
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  var emailCol = customerHeaders.indexOf("Email Address");
  var phoneCol = customerHeaders.indexOf("Phone Number");
  var firstNameCol = customerHeaders.indexOf("First Name");
  var lastNameCol = customerHeaders.indexOf("Last Name");
  var firstVisitCol = customerHeaders.indexOf("First Visit");
  
  var matchedEnvisionCustomers = new Set();
  var customerIdToEnvisionKey = {};
  
  for (var i = 1; i < customerData.length; i++) {
    var customerId = customerData[i][0];
    var firstVisit = customerData[i][firstVisitCol];
    
    if (!isDateInRange(firstVisit)) {
      continue;
    }
    
    var email = normalizeEmail(customerData[i][emailCol]);
    var phone = normalizePhone(customerData[i][phoneCol]);
    var firstName = normalizeString(customerData[i][firstNameCol]);
    var lastName = normalizeString(customerData[i][lastNameCol]);
    
    var matchedKey = null;
    
    if (email && envisionByEmail[email]) {
      matchedKey = envisionByEmail[email];
    } else if (phone && envisionByPhone[phone]) {
      matchedKey = envisionByPhone[phone];
    } else if (firstName && lastName) {
      var nameKey = firstName + "|" + lastName;
      if (envisionByName[nameKey]) {
        matchedKey = envisionByName[nameKey];
      }
    }
    
    if (matchedKey) {
      matchedEnvisionCustomers.add(matchedKey);
      customerIdToEnvisionKey[customerId] = matchedKey;
    }
  }
  
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  var customerIdCol = transHeaders.indexOf("Customer ID");
  var collectedCol = transHeaders.indexOf("Total Collected");
  var transDateCol = transHeaders.indexOf("Date");
  
  var envisionCustomerRevenue = 0;
  
  for (var i = 1; i < transData.length; i++) {
    var transDate = transData[i][transDateCol];
    
    if (!isDateInRange(transDate)) {
      continue;
    }
    
    var customerId = transData[i][customerIdCol];
    var revenue = parseFloat(transData[i][collectedCol]) || 0;
    
    if (customerIdToEnvisionKey[customerId]) {
      envisionCustomerRevenue += revenue;
    }
  }
  
  var totalEnvisionCustomers = envisionData.length;
  
  return {
    totalEnvision: totalEnvisionCustomers,
    signupsFromEnvision: matchedEnvisionSignups.size,
    signupRetentionRate: (matchedEnvisionSignups.size / totalEnvisionCustomers * 100).toFixed(1),
    customersFromEnvision: matchedEnvisionCustomers.size,
    customerRetentionRate: (matchedEnvisionCustomers.size / totalEnvisionCustomers * 100).toFixed(1),
    revenueFromEnvision: envisionCustomerRevenue.toFixed(2)
  };
}

/**
 * Extract revenue metrics for dashboard
 */
function getRevenueMetrics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  
  var collectedCol = transHeaders.indexOf("Total Collected");
  var netTotalCol = transHeaders.indexOf("Net Total");
  var tipCol = transHeaders.indexOf("Tip");
  var customerIdCol = transHeaders.indexOf("Customer ID");
  var transIdCol = transHeaders.indexOf("Transaction ID");
  var transDateCol = transHeaders.indexOf("Date");
  
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  var itemNameCol = itemHeaders.indexOf("Item");
  var itemSalesCol = itemHeaders.indexOf("Gross Sales");

  var eventTransactions = {};
  var eventRevenue = 0;

  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    var category = itemData[i][itemCategoryCol];
    var itemName = itemData[i][itemNameCol];
    var sales = parseFloat(itemData[i][itemSalesCol]) || 0;

    // Use getMajorCategory for consistent event detection
    var majorCat = getMajorCategory(category, itemName);
    if (majorCat === "Event") {
      eventTransactions[transId] = true;
      eventRevenue += sales;
    }
  }
  
  var totalRevenue = 0;
  var totalNetRevenue = 0;
  var totalTips = 0;
  var totalRevenueExcludingEvents = 0;
  var transCountExcludingEvents = 0;
  var customerSpendingExcludingEvents = {};
  
  for (var i = 1; i < transData.length; i++) {
    var transDate = transData[i][transDateCol];
    
    if (!isDateInRange(transDate)) {
      continue;
    }
    
    var collected = parseFloat(transData[i][collectedCol]) || 0;
    var netTotal = parseFloat(transData[i][netTotalCol]) || 0;
    var tip = parseFloat(transData[i][tipCol]) || 0;
    var customerId = transData[i][customerIdCol];
    var transId = transData[i][transIdCol];
    
    totalRevenue += collected;
    totalNetRevenue += netTotal;
    totalTips += tip;
    
    if (!eventTransactions[transId]) {
      totalRevenueExcludingEvents += collected;
      transCountExcludingEvents++;
      
      if (customerId) {
        customerSpendingExcludingEvents[customerId] = (customerSpendingExcludingEvents[customerId] || 0) + collected;
      }
    }
  }
  
  var avgSpendPerVisit = transCountExcludingEvents > 0 ? totalRevenueExcludingEvents / transCountExcludingEvents : 0;
  
  var customerLTVs = [];
  for (var customerId in customerSpendingExcludingEvents) {
    customerLTVs.push(customerSpendingExcludingEvents[customerId]);
  }
  var avgCustomerLTV = customerLTVs.length > 0 ? 
    customerLTVs.reduce(function(a, b) { return a + b; }, 0) / customerLTVs.length : 0;
  
  return {
    totalRevenue: totalRevenue.toFixed(2),
    totalNetRevenue: totalNetRevenue.toFixed(2),
    totalTips: totalTips.toFixed(2),
    totalRevenueExclEvents: totalRevenueExcludingEvents.toFixed(2),
    eventRevenue: eventRevenue.toFixed(2),
    transactionsExclEvents: transCountExcludingEvents,
    avgSpend: avgSpendPerVisit.toFixed(2),
    avgLTV: avgCustomerLTV.toFixed(2)
  };
}

/**
 * Extract category metrics for dashboard
 */
function getCategoryMetrics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var categoryCol = itemHeaders.indexOf("Category");
  var itemNameCol = itemHeaders.indexOf("Item");
  var grossSalesCol = itemHeaders.indexOf("Gross Sales");
  var transIdCol = itemHeaders.indexOf("Transaction ID");
  
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  var collectedCol = transHeaders.indexOf("Total Collected");
  var transactionIdCol = transHeaders.indexOf("Transaction ID");
  var transDateCol = transHeaders.indexOf("Date");
  
  // Build date map for transactions
  var transactionDates = {};
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transactionIdCol];
    var date = transData[i][transDateCol];
    transactionDates[transId] = date;
  }
  
  var eventTransactions = {};
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][transIdCol];
    var category = itemData[i][categoryCol];
    if (category && String(category).toLowerCase().trim() === "event") {
      eventTransactions[transId] = true;
    }
  }
  
  var totalRevenueExcludingEvents = 0;
  var transCountExcludingEvents = 0;
  
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transactionIdCol];
    var revenue = parseFloat(transData[i][collectedCol]) || 0;
    var date = transData[i][transDateCol];
    
    if (!isDateInRange(date)) {
      continue;
    }
    
    if (!eventTransactions[transId]) {
      totalRevenueExcludingEvents += revenue;
      transCountExcludingEvents++;
    }
  }
  
  var majorCategories = {
    "Food": 0,
    "Beverage": 0,
    "Golf": 0,
    "Membership": 0,
    "Miscellaneous": 0
  };
  
  var transactionsWithFB = new Set();
  
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][transIdCol];
    
    if (!isDateInRange(transactionDates[transId])) {
      continue;
    }
    
    var category = itemData[i][categoryCol] || "Uncategorized";
    var itemName = itemData[i][itemNameCol];
    var sales = parseFloat(itemData[i][grossSalesCol]) || 0;
    
    var majorCategory = getMajorCategory(category, itemName);
    
    if (majorCategory !== "Event") {
      majorCategories[majorCategory] += sales;
    }
    
    if (!eventTransactions[transId]) {
      if (majorCategory === "Food" || majorCategory === "Beverage") {
        transactionsWithFB.add(transId);
      }
    }
  }
  
  var totalFBRevenue = majorCategories["Food"] + majorCategories["Beverage"];
  var fbPercent = totalRevenueExcludingEvents > 0 ? (totalFBRevenue / totalRevenueExcludingEvents * 100).toFixed(1) : "0.0";
  var avgFBPerTransaction = transCountExcludingEvents > 0 ? totalFBRevenue / transCountExcludingEvents : 0;
  var fbAttachRate = transCountExcludingEvents > 0 ? (transactionsWithFB.size / transCountExcludingEvents * 100).toFixed(1) : "0.0";
  
  return {
    foodRevenue: majorCategories["Food"].toFixed(2),
    beverageRevenue: majorCategories["Beverage"].toFixed(2),
    golfRevenue: majorCategories["Golf"].toFixed(2),
    membershipRevenue: majorCategories["Membership"].toFixed(2),
    miscRevenue: majorCategories["Miscellaneous"].toFixed(2),
    foodPercent: totalRevenueExcludingEvents > 0 ? (majorCategories["Food"] / totalRevenueExcludingEvents * 100).toFixed(1) : "0.0",
    beveragePercent: totalRevenueExcludingEvents > 0 ? (majorCategories["Beverage"] / totalRevenueExcludingEvents * 100).toFixed(1) : "0.0",
    golfPercent: totalRevenueExcludingEvents > 0 ? (majorCategories["Golf"] / totalRevenueExcludingEvents * 100).toFixed(1) : "0.0",
    membershipPercent: totalRevenueExcludingEvents > 0 ? (majorCategories["Membership"] / totalRevenueExcludingEvents * 100).toFixed(1) : "0.0",
    miscPercent: totalRevenueExcludingEvents > 0 ? (majorCategories["Miscellaneous"] / totalRevenueExcludingEvents * 100).toFixed(1) : "0.0",
    totalFBRevenue: totalFBRevenue.toFixed(2),
    fbPercent: fbPercent,
    avgFBPerTrans: avgFBPerTransaction.toFixed(2),
    fbAttachRate: fbAttachRate
  };
}

/**
 * Extract booking metrics for dashboard
 */
function getBookingMetrics() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var bookingSheet = ss.getSheetByName("Apex Bookings Export");
  var bookingData = bookingSheet.getDataRange().getValues();
  var bookingHeaders = bookingData[0];
  
  var timeCol = bookingHeaders.indexOf("Time");
  var dateCol = bookingHeaders.indexOf("Date");
  var durationCol = bookingHeaders.indexOf("Duration Mins");
  
  var hourCounts = {};
  var dayCounts = {0: 0, 1: 0, 2: 0, 3: 0, 4: 0, 5: 0, 6: 0};
  var durations = [];
  var totalBookings = 0;
  
  for (var i = 1; i < bookingData.length; i++) {
    var date = bookingData[i][dateCol];
    
    if (!isDateInRange(date)) {
      continue;
    }
    
    totalBookings++;
    
    var time = bookingData[i][timeCol];
    var duration = parseFloat(bookingData[i][durationCol]) || 0;
    
    if (time) {
      var hour = typeof time === 'string' ? parseInt(time.split(':')[0]) : time.getHours();
      hourCounts[hour] = (hourCounts[hour] || 0) + 1;
    }
    
    if (date && date instanceof Date) {
      var day = date.getDay();
      dayCounts[day]++;
    }
    
    if (duration > 0) {
      durations.push(duration);
    }
  }
  
  var peakHour = Object.keys(hourCounts).length > 0 ? Object.keys(hourCounts).reduce(function(a, b) {
    return hourCounts[a] > hourCounts[b] ? a : b;
  }) : "N/A";
  
  var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  var mostPopularDay = Object.keys(dayCounts).reduce(function(a, b) {
    return dayCounts[a] > dayCounts[b] ? a : b;
  });
  
  var avgDuration = durations.length > 0 ? 
    durations.reduce(function(a,b){return a+b;},0) / durations.length : 0;
  
  return {
    totalBookings: totalBookings,
    peakHour: peakHour !== "N/A" ? peakHour + ":00" : "N/A",
    popularDay: dayNames[mostPopularDay],
    avgDuration: avgDuration.toFixed(0)
  };
}

// [Continue to File 3 for remaining functions...]