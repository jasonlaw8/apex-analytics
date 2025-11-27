/**
 * ========================================
 * APEX GOLF ANALYTICS - FILE 1 OF 4
 * ========================================
 * 
 * THIS FILE CONTAINS:
 * - Global configuration and date range variables
 * - Custom menu setup
 * - Date range filtering functions
 * - NEW: Misc category analyzer
 * - NEW: Day vs Night spending analysis
 * 
 * DEPENDENCIES: This file calls functions from Files 2, 3, and 4
 */

// ============================================
// DATE RANGE MANAGEMENT (using PropertiesService for persistence)
// ============================================

/**
 * Get current date range settings
 */
function getDateRangeSettings() {
  var props = PropertiesService.getUserProperties();
  
  var startDateStr = props.getProperty('ANALYSIS_START_DATE');
  var endDateStr = props.getProperty('ANALYSIS_END_DATE');
  var label = props.getProperty('ANALYSIS_RANGE_LABEL');
  
  return {
    startDate: startDateStr ? new Date(startDateStr) : null,
    endDate: endDateStr ? new Date(endDateStr) : null,
    label: label || "All Time"
  };
}

/**
 * Save date range settings
 */
function saveDateRangeSettings(startDate, endDate, label) {
  var props = PropertiesService.getUserProperties();
  
  if (startDate) {
    props.setProperty('ANALYSIS_START_DATE', startDate.toISOString());
  } else {
    props.deleteProperty('ANALYSIS_START_DATE');
  }
  
  if (endDate) {
    props.setProperty('ANALYSIS_END_DATE', endDate.toISOString());
  } else {
    props.deleteProperty('ANALYSIS_END_DATE');
  }
  
  props.setProperty('ANALYSIS_RANGE_LABEL', label);
}

// Backward compatibility - these now read from PropertiesService
var ANALYSIS_START_DATE = null;
var ANALYSIS_END_DATE = null;
var ANALYSIS_RANGE_LABEL = "All Time";

// ============================================
// CUSTOM MENU
// ============================================

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('📊 Apex Analytics')
      .addItem('🚀 Run Complete Analysis', 'runCompleteAnalysis')
      .addSeparator()
      .addSubMenu(ui.createMenu('📅 Set Date Range')
          .addItem('All Time', 'setDateRangeAllTime')
          .addItem('Last 7 Days', 'setDateRangeLast7Days')
          .addItem('Last 30 Days', 'setDateRangeLast30Days')
          .addItem('Last 90 Days', 'setDateRangeLast90Days')
          .addItem('Custom Range...', 'setDateRangeCustom'))
      .addSeparator()
      .addItem('👥 Customer Analysis Only', 'runCustomerAnalysisOnly')
      .addItem('💰 Revenue Analysis Only', 'runRevenueAnalysisOnly')
      .addItem('🔄 Envision Retention Only', 'runEnvisionAnalysisOnly')
      .addItem('📅 Booking Analysis Only', 'runBookingAnalysisOnly')
      .addItem('🍔 F&B Analysis Only', 'runFoodBevAnalysisOnly')
      .addItem('🔥 Bonus Insights Only', 'runBonusInsightsOnly')
      .addSeparator()
      .addSubMenu(ui.createMenu('🔍 Special Analysis')
          .addItem('📦 Analyze Misc Items', 'analyzeMiscItems')
          .addItem('⏰ Day vs Night Spending', 'analyzeDayVsNight'))
      .addSeparator()
      .addItem('🧹 Clean Data (Categorize Items)', 'runDataCleanup')
      .addItem('📋 View Uncategorized Items', 'viewUncategorizedItems')
      .addSeparator()
      .addItem('📌 Setup Category Override', 'setupCategoryOverride')
      .addItem('👁️ View Category Overrides', 'viewCategoryOverrides')
      .addSeparator()
      .addItem('🔧 Setup Item Transaction Override', 'setupItemTransactionOverride')
      .addItem('👁️ View Transaction Overrides', 'viewItemTransactionOverrides')
      .addSeparator()
      .addItem('🚨 Validate Bookings vs Payments', 'validateBookingsVsPayments')
      .addSeparator()
      .addItem('ℹ️ About', 'showAbout')
      .addItem('🔍 Diagnostic Check', 'diagnosticCheck')
            .addSeparator()
      .addSubMenu(ui.createMenu('👤 Membership Analytics')
    .addItem('🔍 Query Member Profile', 'queryMemberProfile')
    .addItem('💎 Find Membership Leads', 'findMembershipLeads'))
          .addSeparator()
    .addItem('💰 Calculate Tip Distribution', 'calculateTipDistribution')
              .addSeparator()
    .addItem('📥 Import New Data', 'showImportMenu')
      .addToUi();
    
}

function showAbout() {
  SpreadsheetApp.getUi().alert('Apex Golf Analytics',
    'Analyzes customer data, revenue, bookings, and F&B performance.\n\n' +
    'Distinguishes between signups (profiles created) and customers (those who have transacted).\n\n' +
    'NEW: Date range filtering, Misc analysis, Day/Night comparison',
    SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================
// DATE RANGE FUNCTIONS
// ============================================

function setDateRangeAllTime() {
  saveDateRangeSettings(null, null, "All Time");
  SpreadsheetApp.getUi().alert('✅ Date Range Set', 'Analyzing: All Time\n\nRun analysis to see results.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setDateRangeLast7Days() {
  var today = new Date();
  today.setHours(23, 59, 59, 999);
  
  var startDate = new Date(today);
  startDate.setDate(startDate.getDate() - 7);
  startDate.setHours(0, 0, 0, 0);
  
  saveDateRangeSettings(startDate, today, "Last 7 Days");
  SpreadsheetApp.getUi().alert('✅ Date Range Set', 'Analyzing: Last 7 Days\n\nRun analysis to see results.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setDateRangeLast30Days() {
  var today = new Date();
  today.setHours(23, 59, 59, 999);
  
  var startDate = new Date(today);
  startDate.setDate(startDate.getDate() - 30);
  startDate.setHours(0, 0, 0, 0);
  
  saveDateRangeSettings(startDate, today, "Last 30 Days");
  SpreadsheetApp.getUi().alert('✅ Date Range Set', 'Analyzing: Last 30 Days\n\nRun analysis to see results.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setDateRangeLast90Days() {
  var today = new Date();
  today.setHours(23, 59, 59, 999);
  
  var startDate = new Date(today);
  startDate.setDate(startDate.getDate() - 90);
  startDate.setHours(0, 0, 0, 0);
  
  saveDateRangeSettings(startDate, today, "Last 90 Days");
  SpreadsheetApp.getUi().alert('✅ Date Range Set', 'Analyzing: Last 90 Days\n\nRun analysis to see results.', SpreadsheetApp.getUi().ButtonSet.OK);
}

function setDateRangeCustom() {
  var ui = SpreadsheetApp.getUi();
  
  var startResponse = ui.prompt('Custom Date Range',
    'Enter START date (MM/DD/YYYY):',
    ui.ButtonSet.OK_CANCEL);
  
  if (startResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var endResponse = ui.prompt('Custom Date Range',
    'Enter END date (MM/DD/YYYY):',
    ui.ButtonSet.OK_CANCEL);
  
  if (endResponse.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  try {
    var startDate = new Date(startResponse.getResponseText());
    var endDate = new Date(endResponse.getResponseText());
    
    if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) {
      ui.alert('❌ Invalid Date', 'Please use MM/DD/YYYY format', ui.ButtonSet.OK);
      return;
    }
    
    if (startDate > endDate) {
      ui.alert('❌ Invalid Range', 'Start date must be before end date', ui.ButtonSet.OK);
      return;
    }
    
    startDate.setHours(0, 0, 0, 0);
    endDate.setHours(23, 59, 59, 999);
    
    var label = Utilities.formatDate(startDate, Session.getScriptTimeZone(), "MM/dd/yyyy") + 
                " - " + 
                Utilities.formatDate(endDate, Session.getScriptTimeZone(), "MM/dd/yyyy");
    
    saveDateRangeSettings(startDate, endDate, label);
    
    ui.alert('✅ Date Range Set', 
      'Analyzing: ' + label + '\n\nRun analysis to see results.', 
      ui.ButtonSet.OK);
  } catch (e) {
    ui.alert('❌ Error', 'Invalid date format. Please use MM/DD/YYYY', ui.ButtonSet.OK);
  }
}

// Cache for date range settings (reset on each script execution)
var _cachedDateSettings = null;

/**
 * Helper function to check if a date is within the analysis range
 * Caches settings to avoid repeated PropertiesService calls
 */
function isDateInRange(date) {
  if (!date) return false;
  
  var dateObj = date instanceof Date ? date : new Date(date);
  
  // Get cached settings or fetch once
  if (!_cachedDateSettings) {
    _cachedDateSettings = getDateRangeSettings();
  }
  
  if (_cachedDateSettings.startDate && dateObj < _cachedDateSettings.startDate) {
    return false;
  }
  
  if (_cachedDateSettings.endDate && dateObj > _cachedDateSettings.endDate) {
    return false;
  }
  
  return true;
}

// ============================================
// MISC CATEGORY ANALYZER
// ============================================

/**
 * Analyzes items categorized as Miscellaneous
 * FIXED: Uses consistent getMajorCategory with all overrides
 */
function analyzeMiscItems() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var dateSettings = getDateRangeSettings();

  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  if (!itemSheet) {
    ui.alert('Error', 'Cannot find "Square Item Detail Export" sheet!', ui.ButtonSet.OK);
    return;
  }

  var data = itemSheet.getDataRange().getValues();
  var headers = data[0];

  var itemNameCol = headers.indexOf("Item");
  var categoryCol = headers.indexOf("Category");
  var grossSalesCol = headers.indexOf("Gross Sales");
  var transIdCol = headers.indexOf("Transaction ID");

  // Get transaction dates for filtering
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  var transDateCol = transHeaders.indexOf("Date");
  var transTransIdCol = transHeaders.indexOf("Transaction ID");

  // Build transaction date map
  var transactionDates = {};
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transTransIdCol];
    var date = transData[i][transDateCol];
    if (transId && date) {
      transactionDates[transId] = date;
    }
  }

  // Load all overrides ONCE
  var overridesMaps = {
    transactionOverrides: getItemTransactionOverridesMap(),
    dataCleanup: getDataCleanupMappings(),
    categoryOverrides: null
  };

  // Track misc items
  var miscItems = {};

  for (var i = 1; i < data.length; i++) {
    var transId = data[i][transIdCol];
    var itemName = data[i][itemNameCol];
    var category = data[i][categoryCol];
    var grossSales = parseFloat(data[i][grossSalesCol]) || 0;

    // Apply date filter
    if (!isDateInRange(transactionDates[transId])) {
      continue;
    }

    // Use consistent getMajorCategory with all overrides
    var majorCategory = getMajorCategory(category, itemName, transId, overridesMaps);
    
    // Only track Misc items
    if (majorCategory === "Miscellaneous") {
      if (!miscItems[itemName]) {
        miscItems[itemName] = {
          revenue: 0,
          count: 0,
          originalCategory: category
        };
      }
      miscItems[itemName].revenue += grossSales;
      miscItems[itemName].count++;
    }
  }
  
  // Convert to array and sort by revenue
  var miscArray = [];
  for (var item in miscItems) {
    miscArray.push({
      name: item,
      revenue: miscItems[item].revenue,
      count: miscItems[item].count,
      originalCategory: miscItems[item].originalCategory
    });
  }
  
  miscArray.sort(function(a, b) {
    return b.revenue - a.revenue;
  });
  
  // Create report sheet
  var reportSheet = ss.getSheetByName("Misc Items Analysis");
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet("Misc Items Analysis");
  }
  
  // Header
  reportSheet.getRange("A1:E1").merge();
  reportSheet.getRange("A1").setValue("📦 MISCELLANEOUS ITEMS BREAKDOWN");
  reportSheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.getRange("A1").setBackground("#9C27B0").setFontColor("white");
  
  reportSheet.getRange("A2:E2").merge();
  reportSheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Items currently categorized as 'Miscellaneous' - Review and recategorize as needed");
  reportSheet.getRange("A2").setFontSize(10).setHorizontalAlignment("center").setBackground("#f3f3f3");
  
  // Column headers
  var colHeaders = ["Item Name", "Original Category", "Times Sold", "Total Revenue", "Suggested Category"];
  reportSheet.getRange(4, 1, 1, colHeaders.length).setValues([colHeaders]);
  reportSheet.getRange(4, 1, 1, colHeaders.length).setFontWeight("bold").setBackground("#E8E8E8");
  reportSheet.setFrozenRows(4);
  
  // Data
  var outputData = [];
  for (var i = 0; i < miscArray.length; i++) {
    var suggested = suggestCategory(miscArray[i].name, miscArray[i].originalCategory);
    
    outputData.push([
      miscArray[i].name,
      miscArray[i].originalCategory,
      miscArray[i].count,
      miscArray[i].revenue,
      suggested
    ]);
  }
  
  if (outputData.length > 0) {
    reportSheet.getRange(5, 1, outputData.length, colHeaders.length).setValues(outputData);
    
    // Format revenue column
    reportSheet.getRange(5, 4, outputData.length, 1).setNumberFormat("$#,##0.00");
    
    // Add borders
    reportSheet.getRange(4, 1, outputData.length + 1, colHeaders.length).setBorder(true, true, true, true, true, true);
    
    // Highlight top items
    if (outputData.length >= 1) reportSheet.getRange(5, 1, 1, colHeaders.length).setBackground("#fff3e0");
    if (outputData.length >= 2) reportSheet.getRange(6, 1, 1, colHeaders.length).setBackground("#ffe0b2");
    if (outputData.length >= 3) reportSheet.getRange(7, 1, 1, colHeaders.length).setBackground("#ffcc80");
  }
  
  // Summary at bottom
  var totalMiscRevenue = miscArray.reduce(function(sum, item) { return sum + item.revenue; }, 0);
  var totalMiscCount = miscArray.reduce(function(sum, item) { return sum + item.count; }, 0);
  
  var summaryRow = 5 + outputData.length + 2;
  reportSheet.getRange(summaryRow, 1, 1, 5).merge();
  reportSheet.getRange(summaryRow, 1).setValue(
    "TOTAL: " + miscArray.length + " unique items | " + 
    totalMiscCount + " transactions | $" + totalMiscRevenue.toFixed(2) + " revenue"
  );
  reportSheet.getRange(summaryRow, 1).setFontWeight("bold").setBackground("#e8f0fe");
  
  // Auto-resize
  for (var i = 1; i <= colHeaders.length; i++) {
    reportSheet.autoResizeColumn(i);
  }
  
  ss.setActiveSheet(reportSheet);
  
  var summary = '✅ Analysis Complete (' + dateSettings.label + ')\n\n' +
    'Found ' + miscArray.length + ' items in Misc category\n\n';
  
  if (miscArray.length > 0) {
    summary += 'Top 3 by revenue:\n';
    if (miscArray[0]) summary += '1. ' + miscArray[0].name + ' - $' + miscArray[0].revenue.toFixed(2) + '\n';
    if (miscArray[1]) summary += '2. ' + miscArray[1].name + ' - $' + miscArray[1].revenue.toFixed(2) + '\n';
    if (miscArray[2]) summary += '3. ' + miscArray[2].name + ' - $' + miscArray[2].revenue.toFixed(2) + '\n';
  }
  
  summary += '\nCheck "Misc Items Analysis" sheet for full list and suggestions.';
  
  ui.alert('Misc Items Analysis', summary, ui.ButtonSet.OK);
}

function suggestCategory(itemName, originalCategory) {
  var name = itemName.toLowerCase();
  
  // Food indicators
  if (name.match(/burger|sandwich|pizza|nachos|wings|fries|salad|chicken|taco|wrap|quesadilla|hot dog|pretzel|appetizer|entree/)) {
    return "→ Food?";
  }
  
  // Beverage indicators
  if (name.match(/beer|wine|cocktail|drink|soda|water|juice|coffee|tea|latte|margarita|mojito|shot|ale|ipa|cider/)) {
    return "→ Beverage?";
  }
  
  // Golf indicators
  if (name.match(/bay|simulator|golf|lesson|rental|range|ball|club|hour|round/)) {
    return "→ Golf?";
  }
  
  // Membership indicators
  if (name.match(/membership|member|monthly|annual|subscription|dues/)) {
    return "→ Membership?";
  }
  
  // Event indicators
  if (name.match(/event|party|tournament|league|group|private|corporate|booking/)) {
    return "→ Event?";
  }
  
  return "❓ Review needed";
}

// ============================================
// DAY VS NIGHT SPENDING ANALYZER
// ============================================

/**
 * Analyzes spending patterns by time of day
 * FIXED: Uses same calculation methods as dashboard for consistency
 * FIXED: Applies all category overrides and Data Cleanup mappings
 */
function analyzeDayVsNight() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();

  var dateSettings = getDateRangeSettings();

  var transSheet = ss.getSheetByName("Square Transactions Export");
  var itemSheet = ss.getSheetByName("Square Item Detail Export");

  if (!transSheet || !itemSheet) {
    ui.alert('Error', 'Missing required sheets: Square Transactions Export, Square Item Detail Export', ui.ButtonSet.OK);
    return;
  }

  var transData = transSheet.getDataRange().getValues();
  var itemData = itemSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  var itemHeaders = itemData[0];

  var transIdCol = transHeaders.indexOf("Transaction ID");
  var transTimeCol = transHeaders.indexOf("Time");
  var transDateCol = transHeaders.indexOf("Date");

  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemNameCol = itemHeaders.indexOf("Item");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  var itemGrossCol = itemHeaders.indexOf("Gross Sales");

  // Load all overrides ONCE for performance
  var overridesMaps = {
    transactionOverrides: getItemTransactionOverridesMap(),
    dataCleanup: getDataCleanupMappings(),
    categoryOverrides: null
  };
  
  // Track 4 time periods
  var periods = {
    morning: {   // 6am - 11:59am
      name: "Morning (6am-12pm)",
      totalRevenue: 0,
      totalRevenueWithGolf: 0,
      fbOnly: 0,
      transCount: 0,
      fbTransCount: 0
    },
    afternoon: { // 12pm - 5:59pm
      name: "Afternoon (12pm-6pm)",
      totalRevenue: 0,
      totalRevenueWithGolf: 0,
      fbOnly: 0,
      transCount: 0,
      fbTransCount: 0
    },
    evening: {   // 6pm - 8:59pm
      name: "Evening (6pm-9pm)",
      totalRevenue: 0,
      totalRevenueWithGolf: 0,
      fbOnly: 0,
      transCount: 0,
      fbTransCount: 0
    },
    night: {     // 9pm - 5:59am
      name: "Night (9pm-6am)",
      totalRevenue: 0,
      totalRevenueWithGolf: 0,
      fbOnly: 0,
      transCount: 0,
      fbTransCount: 0
    }
  };
  
  // Build transaction time map and categorize items
  var transactionCategories = {};
  
  // Build transaction date lookup
  var transactionDates = {};
  var transactionTimes = {};
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transIdCol];
    transactionDates[transId] = transData[i][transDateCol];
    transactionTimes[transId] = transData[i][transTimeCol];
  }

  // First pass: categorize all items using consistent method
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    var category = itemData[i][itemCategoryCol];
    var itemName = itemData[i][itemNameCol];
    var grossSales = parseFloat(itemData[i][itemGrossCol]) || 0;

    // Apply date filter
    if (!isDateInRange(transactionDates[transId])) {
      continue;
    }

    // Use consistent getMajorCategory with all overrides
    var majorCat = getMajorCategory(category, itemName, transId, overridesMaps);
    
    // Skip Events and Memberships
    if (majorCat === "Event" || majorCat === "Membership") {
      continue;
    }
    
    if (!transactionCategories[transId]) {
      transactionCategories[transId] = {
        hasGolf: false,
        hasFB: false,
        fbAmount: 0,
        totalAmount: 0
      };
    }
    
    transactionCategories[transId].totalAmount += grossSales;
    
    if (majorCat === "Golf") {
      transactionCategories[transId].hasGolf = true;
    } else if (majorCat === "Food" || majorCat === "Beverage") {
      transactionCategories[transId].hasFB = true;
      transactionCategories[transId].fbAmount += grossSales;
    }
  }
  
  // Second pass: analyze by time period
  for (var transId in transactionCategories) {
    var time = transactionTimes[transId];
    var date = transactionDates[transId];

    // Apply date filter
    if (!isDateInRange(date)) {
      continue;
    }

    // Determine time period
    var hour = getHourFromTime(time);
    if (hour === null) continue;
    
    var period;
    if (hour >= 6 && hour < 12) {
      period = periods.morning;
    } else if (hour >= 12 && hour < 18) {
      period = periods.afternoon;
    } else if (hour >= 18 && hour < 21) {
      period = periods.evening;
    } else {
      period = periods.night;
    }
    
    var catData = transactionCategories[transId];
    
    // Total with golf
    period.totalRevenueWithGolf += catData.totalAmount;
    period.transCount++;
    
    // F&B only transactions
    if (catData.hasFB && !catData.hasGolf) {
      period.fbOnly += catData.fbAmount;
      period.fbTransCount++;
    }
    
    // Total without golf
    if (!catData.hasGolf && catData.totalAmount > 0) {
      period.totalRevenue += catData.totalAmount;
    }
  }
  
  // Create report
  var reportSheet = ss.getSheetByName("Time Period Analysis");
  if (reportSheet) {
    reportSheet.clear();
  } else {
    reportSheet = ss.insertSheet("Time Period Analysis");
  }
  
  // Header
  reportSheet.getRange("A1:F1").merge();
  reportSheet.getRange("A1").setValue("⏰ SPENDING BY TIME OF DAY");
  reportSheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  reportSheet.getRange("A1").setBackground("#FF6D00").setFontColor("white");
  
  reportSheet.getRange("A2:F2").merge();
  reportSheet.getRange("A2").setValue("Date Range: " + dateSettings.label + " | Morning (6am-12pm) | Afternoon (12pm-6pm) | Evening (6pm-9pm) | Night (9pm-6am) | Excludes Events & Memberships");
  reportSheet.getRange("A2").setFontSize(9).setHorizontalAlignment("center").setBackground("#f3f3f3");
  
  var currentRow = 4;
  
  // WITH GOLF SECTION
  reportSheet.getRange(currentRow, 1, 1, 6).merge();
  reportSheet.getRange(currentRow, 1).setValue("💰 AVERAGE CHECK SIZE (With Golf - All Spending)");
  reportSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#fff3e0");
  currentRow++;
  
  reportSheet.getRange(currentRow, 1, 1, 6).setValues([["Time Period", "Avg Check", "Transactions", "Total Revenue", "% of Day", ""]]);
  reportSheet.getRange(currentRow, 1, 1, 6).setFontWeight("bold").setBackground("#E8E8E8");
  currentRow++;
  
  var withGolfData = [];
  var periodKeys = ['morning', 'afternoon', 'evening', 'night'];
  var highestAvg = 0;
  var highestPeriod = '';
  var totalTrans = 0;
  var totalRev = 0;
  
  for (var i = 0; i < periodKeys.length; i++) {
    var p = periods[periodKeys[i]];
    var avg = p.transCount > 0 ? p.totalRevenueWithGolf / p.transCount : 0;
    if (avg > highestAvg) {
      highestAvg = avg;
      highestPeriod = p.name;
    }
    totalTrans += p.transCount;
    totalRev += p.totalRevenueWithGolf;
    
    withGolfData.push([
      p.name,
      "$" + avg.toFixed(2),
      p.transCount,
      "$" + p.totalRevenueWithGolf.toFixed(2),
      "", // Will calculate % below
      ""
    ]);
  }
  
  // Calculate percentages
  for (var i = 0; i < withGolfData.length; i++) {
    var pct = totalTrans > 0 ? (withGolfData[i][2] / totalTrans * 100) : 0;
    withGolfData[i][4] = pct.toFixed(1) + "%";
  }
  
  reportSheet.getRange(currentRow, 1, withGolfData.length, 6).setValues(withGolfData);
  reportSheet.getRange(currentRow, 1, withGolfData.length, 6).setBorder(true, true, true, true, true, true);
  
  // Highlight highest
  for (var i = 0; i < withGolfData.length; i++) {
    if (withGolfData[i][0] === highestPeriod) {
      reportSheet.getRange(currentRow + i, 1, 1, 6).setBackground("#d9ead3");
    }
  }
  
  currentRow += withGolfData.length + 2;
  
  // F&B ONLY SECTION
  reportSheet.getRange(currentRow, 1, 1, 6).merge();
  reportSheet.getRange(currentRow, 1).setValue("🍔 AVERAGE CHECK SIZE (F&B Only - No Golf)");
  reportSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#fff3e0");
  currentRow++;
  
  reportSheet.getRange(currentRow, 1, 1, 6).setValues([["Time Period", "Avg Check", "Transactions", "Total F&B", "% of Day", ""]]);
  reportSheet.getRange(currentRow, 1, 1, 6).setFontWeight("bold").setBackground("#E8E8E8");
  currentRow++;
  
  var fbData = [];
  var highestFBAvg = 0;
  var highestFBPeriod = '';
  var totalFBTrans = 0;
  
  for (var i = 0; i < periodKeys.length; i++) {
    var p = periods[periodKeys[i]];
    var avg = p.fbTransCount > 0 ? p.fbOnly / p.fbTransCount : 0;
    if (avg > highestFBAvg) {
      highestFBAvg = avg;
      highestFBPeriod = p.name;
    }
    totalFBTrans += p.fbTransCount;
    
    fbData.push([
      p.name,
      "$" + avg.toFixed(2),
      p.fbTransCount,
      "$" + p.fbOnly.toFixed(2),
      "",
      ""
    ]);
  }
  
  // Calculate percentages
  for (var i = 0; i < fbData.length; i++) {
    var pct = totalFBTrans > 0 ? (fbData[i][2] / totalFBTrans * 100) : 0;
    fbData[i][4] = pct.toFixed(1) + "%";
  }
  
  reportSheet.getRange(currentRow, 1, fbData.length, 6).setValues(fbData);
  reportSheet.getRange(currentRow, 1, fbData.length, 6).setBorder(true, true, true, true, true, true);
  
  // Highlight highest
  for (var i = 0; i < fbData.length; i++) {
    if (fbData[i][0] === highestFBPeriod) {
      reportSheet.getRange(currentRow + i, 1, 1, 6).setBackground("#d9ead3");
    }
  }
  
  currentRow += fbData.length + 2;
  
  // KEY INSIGHTS
  reportSheet.getRange(currentRow, 1, 1, 6).merge();
  reportSheet.getRange(currentRow, 1).setValue("💡 KEY INSIGHTS");
  reportSheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#e8f0fe");
  currentRow++;
  
  var insights = [];
  insights.push(["✓ Highest avg check (with golf): " + highestPeriod + " at $" + highestAvg.toFixed(2), "", "", "", "", ""]);
  insights.push(["✓ Highest F&B spending: " + highestFBPeriod + " at $" + highestFBAvg.toFixed(2), "", "", "", "", ""]);
  
  // Business recommendations
  if (highestPeriod.indexOf("Evening") >= 0 || highestPeriod.indexOf("Night") >= 0) {
    insights.push(["→ Consider premium pricing for evening/night slots", "", "", "", "", ""]);
  }
  if (highestFBPeriod.indexOf("Evening") >= 0 || highestFBPeriod.indexOf("Night") >= 0) {
    insights.push(["→ Promote dinner menu and happy hour specials", "", "", "", "", ""]);
  }
  if (highestPeriod.indexOf("Morning") >= 0) {
    insights.push(["→ Morning slots are premium - maintain pricing", "", "", "", "", ""]);
  }
  if (highestPeriod.indexOf("Afternoon") >= 0) {
    insights.push(["→ Afternoon is peak time - optimize staffing", "", "", "", "", ""]);
  }
  
  reportSheet.getRange(currentRow, 1, insights.length, 6).setValues(insights);
  
  // Auto-resize
  for (var i = 1; i <= 6; i++) {
    reportSheet.autoResizeColumn(i);
  }
  
  ss.setActiveSheet(reportSheet);
  
  // Build summary message
  var summary = 'Time Period Analysis Complete!\n\n';
  summary += 'Date Range: ' + dateSettings.label + '\n\n';
  summary += 'HIGHEST AVG CHECK (with golf):\n' + highestPeriod + ' - $' + highestAvg.toFixed(2) + '\n\n';
  summary += 'HIGHEST F&B SPENDING:\n' + highestFBPeriod + ' - $' + highestFBAvg.toFixed(2) + '\n\n';
  summary += 'Check "Time Period Analysis" sheet for details!';
  
  ui.alert('✅ Analysis Complete', summary, ui.ButtonSet.OK);
}

function getHourFromTime(time) {
  if (!time) return null;
  
  if (typeof time === 'string') {
    // Format like "14:30:00" or "2:30 PM"
    var parts = time.split(':');
    if (parts.length > 0) {
      var hour = parseInt(parts[0]);
      if (!isNaN(hour)) {
        return hour;
      }
    }
  } else if (time instanceof Date) {
    return time.getHours();
  }
  
  return null;
}

// ============================================
// DIAGNOSTIC CHECK
// ============================================

function diagnosticCheck() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  var requiredSheets = [
    "Square Transactions Export",
    "Square Item Detail Export",
    "Square Customer Export",
    "Apex Bookings Export",
    "Customer List",
    "Envision Customer List"
  ];
  
  var report = "DIAGNOSTIC REPORT\n\n";
  var allGood = true;
  
  for (var i = 0; i < requiredSheets.length; i++) {
    var sheetName = requiredSheets[i];
    var sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) {
      report += "❌ MISSING: " + sheetName + "\n";
      allGood = false;
    } else {
      var rowCount = sheet.getLastRow();
      var colCount = sheet.getLastColumn();
      report += "✅ FOUND: " + sheetName + " (" + rowCount + " rows, " + colCount + " cols)\n";
      
      if (rowCount < 2) {
        report += "   ⚠️ WARNING: Sheet appears empty\n";
        allGood = false;
      }
    }
  }
  
  report += "\n";
  
  if (allGood) {
    report += "✅ ALL CHECKS PASSED!\nReady to run analysis.";
  } else {
    report += "❌ ISSUES FOUND\nPlease fix issues above.";
  }
  
  ui.alert("Sheet Diagnostic", report, ui.ButtonSet.OK);
}