/**
 * ========================================
 * APEX GOLF ANALYTICS - FILE 5: MEMBERSHIP
 * ========================================
 * 
 * THIS FILE CONTAINS:
 * - Member Profile Query (detailed individual analysis)
 * - Membership Lead Finder (identify high-potential non-members)
 * 
 * Add to menu in File 1 onOpen():
 * .addSubMenu(ui.createMenu('👤 Membership Analytics')
 *     .addItem('🔍 Query Member Profile', 'queryMemberProfile')
 *     .addItem('💎 Find Membership Leads', 'findMembershipLeads'))
 */

// ============================================
// FUNCTION 1: QUERY MEMBER PROFILE
// ============================================

/**
 * Main function - prompts user to search for a member
 */
function queryMemberProfile() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  // Get customer list
  var customerSheet = ss.getSheetByName("Square Customer Export");
  if (!customerSheet) {
    ui.alert('Error', 'Cannot find "Square Customer Export" sheet!', ui.ButtonSet.OK);
    return;
  }
  
  // Prompt for member search
  var response = ui.prompt(
    'Query Member Profile',
    'Enter member email or name to search:',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() !== ui.Button.OK) {
    return;
  }
  
  var searchTerm = response.getResponseText().toLowerCase().trim();
  
  if (!searchTerm) {
    ui.alert('Error', 'Please enter a search term', ui.ButtonSet.OK);
    return;
  }
  
  // Find matching customer
  var member = findMemberBySearch(searchTerm);
  
  if (!member) {
    ui.alert('Not Found', 'No member found matching: "' + searchTerm + '"', ui.ButtonSet.OK);
    return;
  }
  
  // Generate the profile
  generateMemberProfile(member);
}

/**
 * Search for a member by email or name
 */
function findMemberBySearch(searchTerm) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  
  var customerIdCol = customerHeaders.indexOf("Square Customer ID");
  var firstNameCol = customerHeaders.indexOf("First Name");
  var lastNameCol = customerHeaders.indexOf("Last Name");
  var emailCol = customerHeaders.indexOf("Email Address");
  
  // Search through customers
  for (var i = 1; i < customerData.length; i++) {
    var customerId = customerData[i][customerIdCol];
    var firstName = String(customerData[i][firstNameCol] || "").toLowerCase();
    var lastName = String(customerData[i][lastNameCol] || "").toLowerCase();
    var email = String(customerData[i][emailCol] || "").toLowerCase();
    var fullName = (firstName + " " + lastName).trim();
    
    // Check for matches
    if (email === searchTerm || 
        fullName === searchTerm || 
        firstName === searchTerm || 
        lastName === searchTerm ||
        email.indexOf(searchTerm) >= 0 ||
        fullName.indexOf(searchTerm) >= 0 ||
        firstName.indexOf(searchTerm) >= 0 ||
        lastName.indexOf(searchTerm) >= 0) {
      
      return {
        id: customerId,
        firstName: customerData[i][firstNameCol] || "",
        lastName: customerData[i][lastNameCol] || "",
        email: customerData[i][emailCol] || "",
        fullName: (customerData[i][firstNameCol] + " " + customerData[i][lastNameCol]).trim() || email
      };
    }
  }
  
  return null;
}

/**
 * Generate complete member profile report
 */
function generateMemberProfile(member) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Get or create profile sheet
  var profileSheet = ss.getSheetByName("Member Profile");
  if (profileSheet) {
    profileSheet.clear();
  } else {
    profileSheet = ss.insertSheet("Member Profile");
  }
  
  // Set up sheet formatting
  profileSheet.setColumnWidth(1, 250);
  profileSheet.setColumnWidth(2, 150);
  profileSheet.setColumnWidth(3, 100);
  profileSheet.setColumnWidth(4, 250);
  profileSheet.setColumnWidth(5, 150);
  profileSheet.setColumnWidth(6, 150);
  
  var currentRow = 1;
  
  // === MAIN HEADER ===
  profileSheet.getRange("A1:F1").merge();
  profileSheet.getRange("A1").setValue("👤 MEMBER PROFILE: " + member.fullName);
  profileSheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  profileSheet.getRange("A1").setBackground("#4285F4").setFontColor("white");
  currentRow++;
  
  profileSheet.getRange("A2:F2").merge();
  profileSheet.getRange("A2").setValue("Email: " + member.email + " | Generated: " + new Date());
  profileSheet.getRange("A2").setFontSize(10).setHorizontalAlignment("center").setBackground("#e8f0fe");
  currentRow += 2;
  
  // Get all member data
  var memberData = getMemberData(member.id, member.email);
  
  // === SECTION 1: PLAY TIME STATISTICS ===
  profileSheet.getRange(currentRow, 1, 1, 6).merge();
  profileSheet.getRange(currentRow, 1).setValue("⏱️ PLAY TIME STATISTICS");
  profileSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#34A853").setFontColor("white");
  currentRow++;
  
  var playTimeData = [
    ["Last 30 Days", memberData.playTime.last30Days.toFixed(1) + " hours", "", "Last 90 Days", memberData.playTime.last90Days.toFixed(1) + " hours", ""],
    ["Last Year", memberData.playTime.lastYear.toFixed(1) + " hours", "", "All Time", memberData.playTime.allTime.toFixed(1) + " hours", ""],
    ["Avg Hours/Visit", memberData.playTime.avgPerVisit.toFixed(1) + " hours", "", "Total Visits", memberData.visits.total, ""]
  ];
  
  profileSheet.getRange(currentRow, 1, playTimeData.length, 6).setValues(playTimeData);
  profileSheet.getRange(currentRow, 1, playTimeData.length, 6).setBorder(true, true, true, true, true, true);
  profileSheet.getRange(currentRow, 1, playTimeData.length, 1).setBackground("#d9ead3");
  profileSheet.getRange(currentRow, 4, playTimeData.length, 1).setBackground("#d9ead3");
  currentRow += playTimeData.length + 2;
  
  // === SECTION 2: VISIT PATTERNS ===
  profileSheet.getRange(currentRow, 1, 1, 6).merge();
  profileSheet.getRange(currentRow, 1).setValue("📅 VISIT PATTERNS");
  profileSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#FBBC04").setFontColor("white");
  currentRow++;
  
  var visitData = [
    ["Total Visits", memberData.visits.total, "", "Visits Last 30 Days", memberData.visits.last30Days, ""],
    ["Visits Last 90 Days", memberData.visits.last90Days, "", "Visits Last Year", memberData.visits.lastYear, ""],
    ["Favorite Day", memberData.patterns.favoriteDay, "", "Favorite Time", memberData.patterns.favoriteTime, ""],
    ["Avg Days Between Visits", memberData.patterns.avgDaysBetween, "", "Most Recent Visit", memberData.visits.mostRecent, ""]
  ];
  
  profileSheet.getRange(currentRow, 1, visitData.length, 6).setValues(visitData);
  profileSheet.getRange(currentRow, 1, visitData.length, 6).setBorder(true, true, true, true, true, true);
  profileSheet.getRange(currentRow, 1, visitData.length, 1).setBackground("#fef7e0");
  profileSheet.getRange(currentRow, 4, visitData.length, 1).setBackground("#fef7e0");
  currentRow += visitData.length + 2;
  
  // === SECTION 3: SPENDING STATISTICS ===
  profileSheet.getRange(currentRow, 1, 1, 6).merge();
  profileSheet.getRange(currentRow, 1).setValue("💰 SPENDING STATISTICS");
  profileSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#EA4335").setFontColor("white");
  currentRow++;
  
  var spendingData = [
    ["Total Lifetime Spend", "$" + memberData.spending.lifetime.toFixed(2), "", "Avg Spend/Visit", "$" + memberData.spending.avgPerVisit.toFixed(2), ""],
    ["Last 30 Days Spend", "$" + memberData.spending.last30Days.toFixed(2), "", "Last 90 Days Spend", "$" + memberData.spending.last90Days.toFixed(2), ""],
    ["Last Year Spend", "$" + memberData.spending.lastYear.toFixed(2), "", "", "", ""]
  ];
  
  profileSheet.getRange(currentRow, 1, spendingData.length, 6).setValues(spendingData);
  profileSheet.getRange(currentRow, 1, spendingData.length, 6).setBorder(true, true, true, true, true, true);
  profileSheet.getRange(currentRow, 1, spendingData.length, 1).setBackground("#f4cccc");
  profileSheet.getRange(currentRow, 4, spendingData.length, 1).setBackground("#f4cccc");
  currentRow += spendingData.length + 2;
  
  // === SECTION 4: CATEGORY BREAKDOWN ===
  profileSheet.getRange(currentRow, 1, 1, 6).merge();
  profileSheet.getRange(currentRow, 1).setValue("📊 SPENDING BY CATEGORY");
  profileSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#9C27B0").setFontColor("white");
  currentRow++;
  
  var categoryData = [
    ["⛳ Golf", "$" + memberData.categories.golf.toFixed(2), (memberData.categories.golfPct * 100).toFixed(1) + "%"],
    ["🍔 Food", "$" + memberData.categories.food.toFixed(2), (memberData.categories.foodPct * 100).toFixed(1) + "%"],
    ["🍺 Beverage", "$" + memberData.categories.beverage.toFixed(2), (memberData.categories.beveragePct * 100).toFixed(1) + "%"],
    ["👤 Membership", "$" + memberData.categories.membership.toFixed(2), (memberData.categories.membershipPct * 100).toFixed(1) + "%"],
    ["📦 Other", "$" + memberData.categories.misc.toFixed(2), (memberData.categories.miscPct * 100).toFixed(1) + "%"]
  ];
  
  profileSheet.getRange(currentRow, 1, categoryData.length, 3).setValues(categoryData);
  profileSheet.getRange(currentRow, 1, categoryData.length, 3).setBorder(true, true, true, true, true, true);
  profileSheet.getRange(currentRow, 1, categoryData.length, 1).setBackground("#f3e5f5");
  currentRow += categoryData.length + 2;
  
  // === SECTION 5: F&B BREAKDOWN ===
  profileSheet.getRange(currentRow, 1, 1, 6).merge();
  profileSheet.getRange(currentRow, 1).setValue("🍔 FOOD & BEVERAGE DETAIL");
  profileSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#FF6D00").setFontColor("white");
  currentRow++;
  
  var fbData = [
    ["Total F&B Spend", "$" + memberData.fb.total.toFixed(2), "", "F&B % of Total", (memberData.fb.percentage * 100).toFixed(1) + "%", ""],
    ["Food Spend", "$" + memberData.fb.food.toFixed(2), "", "Beverage Spend", "$" + memberData.fb.beverage.toFixed(2), ""],
    ["Avg F&B/Visit", "$" + memberData.fb.avgPerVisit.toFixed(2), "", "Visits with F&B", memberData.fb.visitsWithFB, ""]
  ];
  
  profileSheet.getRange(currentRow, 1, fbData.length, 6).setValues(fbData);
  profileSheet.getRange(currentRow, 1, fbData.length, 6).setBorder(true, true, true, true, true, true);
  profileSheet.getRange(currentRow, 1, fbData.length, 1).setBackground("#ffe0b2");
  profileSheet.getRange(currentRow, 4, fbData.length, 1).setBackground("#ffe0b2");
  currentRow += fbData.length + 2;
  
  // === SECTION 6: MEMBER INSIGHTS ===
  profileSheet.getRange(currentRow, 1, 1, 6).merge();
  profileSheet.getRange(currentRow, 1).setValue("💡 MEMBER INSIGHTS");
  profileSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#00ACC1").setFontColor("white");
  currentRow++;
  
  var insights = generateMemberInsights(memberData);
  for (var i = 0; i < insights.length; i++) {
    profileSheet.getRange(currentRow, 1, 1, 6).merge();
    profileSheet.getRange(currentRow, 1).setValue("• " + insights[i]);
    profileSheet.getRange(currentRow, 1).setWrap(true);
    currentRow++;
  }
  currentRow += 1;
  
  // === SECTION 7: TRANSACTION HISTORY ===
  profileSheet.getRange(currentRow, 1, 1, 6).merge();
  profileSheet.getRange(currentRow, 1).setValue("📝 COMPLETE TRANSACTION HISTORY");
  profileSheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#607D8B").setFontColor("white");
  currentRow++;
  
  var transHeaders = ["Date", "Items Purchased", "Categories", "Amount", "Duration", "Notes"];
  profileSheet.getRange(currentRow, 1, 1, transHeaders.length).setValues([transHeaders]);
  profileSheet.getRange(currentRow, 1, 1, transHeaders.length).setFontWeight("bold").setBackground("#E8E8E8");
  currentRow++;
  
  if (memberData.transactions.length > 0) {
    profileSheet.getRange(currentRow, 1, memberData.transactions.length, transHeaders.length).setValues(memberData.transactions);
    profileSheet.getRange(currentRow, 1, memberData.transactions.length, transHeaders.length).setBorder(true, true, true, true, true, true);
    profileSheet.getRange(currentRow, 4, memberData.transactions.length, 1).setNumberFormat("$#,##0.00");
  } else {
    profileSheet.getRange(currentRow, 1, 1, 6).merge();
    profileSheet.getRange(currentRow, 1).setValue("No transactions found");
  }
  
  // Activate the sheet
  ss.setActiveSheet(profileSheet);
  
  // Show completion message
  SpreadsheetApp.getUi().alert(
    '✅ Profile Generated!',
    'Member profile for ' + member.fullName + ' is ready!\n\nCheck the "Member Profile" sheet.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Gather all data for a specific member
 */
function getMemberData(customerId, customerEmail) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var now = new Date();
  var thirtyDaysAgo = new Date(now.getTime() - (30 * 24 * 60 * 60 * 1000));
  var ninetyDaysAgo = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));
  var oneYearAgo = new Date(now.getTime() - (365 * 24 * 60 * 60 * 1000));
  
  // Initialize result object
  var result = {
    playTime: {
      last30Days: 0,
      last90Days: 0,
      lastYear: 0,
      allTime: 0,
      avgPerVisit: 0
    },
    visits: {
      total: 0,
      last30Days: 0,
      last90Days: 0,
      lastYear: 0,
      dates: [],
      mostRecent: "N/A"
    },
    spending: {
      lifetime: 0,
      last30Days: 0,
      last90Days: 0,
      lastYear: 0,
      avgPerVisit: 0
    },
    fb: {
      total: 0,
      food: 0,
      beverage: 0,
      percentage: 0,
      avgPerVisit: 0,
      visitsWithFB: 0
    },
    categories: {
      golf: 0,
      food: 0,
      beverage: 0,
      membership: 0,
      misc: 0,
      golfPct: 0,
      foodPct: 0,
      beveragePct: 0,
      membershipPct: 0,
      miscPct: 0
    },
    patterns: {
      favoriteDay: "N/A",
      favoriteTime: "N/A",
      avgDaysBetween: "N/A"
    },
    transactions: []
  };
  
  // === GET PLAY TIME FROM BOOKINGS AND TRACK BOOKING VISITS ===
  var bookingSheet = ss.getSheetByName("Apex Bookings Export");
  var bookingsByDate = {};
  var visitDateSet = {}; // Track ALL unique visit dates (bookings + transactions)

  if (bookingSheet) {
    var bookingData = bookingSheet.getDataRange().getValues();
    var bookingHeaders = bookingData[0];

    var bookingEmailCol = bookingHeaders.indexOf("Email");
    var bookingDateCol = bookingHeaders.indexOf("Date");
    var bookingDurationCol = bookingHeaders.indexOf("Duration Mins");

    var normEmail = normalizeEmail(customerEmail);

    for (var i = 1; i < bookingData.length; i++) {
      var email = normalizeEmail(bookingData[i][bookingEmailCol]);

      if (email === normEmail) {
        var date = bookingData[i][bookingDateCol];
        var duration = parseFloat(bookingData[i][bookingDurationCol]) || 0;
        var hours = duration / 60;

        var dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");
        bookingsByDate[dateStr] = (bookingsByDate[dateStr] || 0) + hours;

        // Track this as a visit date
        if (!visitDateSet[dateStr]) {
          visitDateSet[dateStr] = new Date(date);
          result.visits.dates.push(new Date(date));
        }

        result.playTime.allTime += hours;

        if (date >= thirtyDaysAgo) result.playTime.last30Days += hours;
        if (date >= ninetyDaysAgo) result.playTime.last90Days += hours;
        if (date >= oneYearAgo) result.playTime.lastYear += hours;
      }
    }
  }
  
  // === GET TRANSACTION DATA ===
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];

  var transCustomerIdCol = transHeaders.indexOf("Customer ID");
  var transEmailCol = transHeaders.indexOf("Customer Email");
  var transDateCol = transHeaders.indexOf("Date");
  var transCollectedCol = transHeaders.indexOf("Total Collected");
  var transIdCol = transHeaders.indexOf("Transaction ID");

  var transactionDetails = {};
  var dayCount = {};
  var hourCount = {};
  var spendingByDate = {}; // Track spending per date

  var normEmail = normalizeEmail(customerEmail);

  for (var i = 1; i < transData.length; i++) {
    var transCustomerId = transData[i][transCustomerIdCol];
    var transEmail = normalizeEmail(transData[i][transEmailCol]);

    // Match by BOTH Customer ID AND email to catch all transactions
    if (transCustomerId === customerId || (normEmail && transEmail === normEmail)) {
      var date = transData[i][transDateCol];
      var amount = parseFloat(transData[i][transCollectedCol]) || 0;
      var transId = transData[i][transIdCol];
      var dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "yyyy-MM-dd");

      // Track unique visit dates (add to the set created from bookings)
      if (!visitDateSet[dateStr]) {
        visitDateSet[dateStr] = new Date(date);
        result.visits.dates.push(new Date(date));
      }

      // Accumulate spending
      result.spending.lifetime += amount;
      spendingByDate[dateStr] = (spendingByDate[dateStr] || 0) + amount;

      if (date >= thirtyDaysAgo) {
        result.spending.last30Days += amount;
      }
      if (date >= ninetyDaysAgo) {
        result.spending.last90Days += amount;
      }
      if (date >= oneYearAgo) {
        result.spending.lastYear += amount;
      }

      if (result.visits.mostRecent === "N/A" || date > new Date(result.visits.mostRecent)) {
        result.visits.mostRecent = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), "MM/dd/yyyy");
      }

      transactionDetails[transId] = {
        date: date,
        amount: amount,
        items: [],
        categories: new Set()
      };
    }
  }

  // Count visits from unique dates
  result.visits.total = Object.keys(visitDateSet).length;

  // Count visits by time period based on unique dates
  for (var dateStr in visitDateSet) {
    var visitDate = visitDateSet[dateStr];

    if (visitDate >= thirtyDaysAgo) {
      result.visits.last30Days++;
    }
    if (visitDate >= ninetyDaysAgo) {
      result.visits.last90Days++;
    }
    if (visitDate >= oneYearAgo) {
      result.visits.lastYear++;
    }

    // Track day of week (based on unique visits, not transactions)
    var dayOfWeek = visitDate.getDay();
    dayCount[dayOfWeek] = (dayCount[dayOfWeek] || 0) + 1;

    // Track hour of day (using first transaction of the day)
    var hour = visitDate.getHours();
    hourCount[hour] = (hourCount[hour] || 0) + 1;
  }
  
  // === GET ITEM DETAILS ===
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemNameCol = itemHeaders.indexOf("Item");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  var itemGrossCol = itemHeaders.indexOf("Gross Sales");
  
  var transactionsWithFB = new Set();
  
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    
    if (transactionDetails[transId]) {
      var itemName = itemData[i][itemNameCol] || "";
      var category = itemData[i][itemCategoryCol] || "";
      var grossSales = parseFloat(itemData[i][itemGrossCol]) || 0;
      
      var majorCat = getMajorCategory(category, itemName);
      
      transactionDetails[transId].items.push(itemName);
      transactionDetails[transId].categories.add(majorCat);
      
      // Track by category
      if (majorCat === "Golf") {
        result.categories.golf += grossSales;
      } else if (majorCat === "Food") {
        result.categories.food += grossSales;
        result.fb.food += grossSales;
        result.fb.total += grossSales;
        transactionsWithFB.add(transId);
      } else if (majorCat === "Beverage") {
        result.categories.beverage += grossSales;
        result.fb.beverage += grossSales;
        result.fb.total += grossSales;
        transactionsWithFB.add(transId);
      } else if (majorCat === "Membership") {
        result.categories.membership += grossSales;
      } else {
        result.categories.misc += grossSales;
      }
    }
  }
  
  result.fb.visitsWithFB = transactionsWithFB.size;
  
  // === CALCULATE PERCENTAGES ===
  if (result.spending.lifetime > 0) {
    result.categories.golfPct = result.categories.golf / result.spending.lifetime;
    result.categories.foodPct = result.categories.food / result.spending.lifetime;
    result.categories.beveragePct = result.categories.beverage / result.spending.lifetime;
    result.categories.membershipPct = result.categories.membership / result.spending.lifetime;
    result.categories.miscPct = result.categories.misc / result.spending.lifetime;
    result.fb.percentage = result.fb.total / result.spending.lifetime;
  }
  
  // === CALCULATE AVERAGES ===
  if (result.visits.total > 0) {
    result.spending.avgPerVisit = result.spending.lifetime / result.visits.total;
    result.fb.avgPerVisit = result.fb.total / result.visits.total;
    result.playTime.avgPerVisit = result.playTime.allTime / result.visits.total;
  }
  
  // === FIND FAVORITE DAY ===
  if (Object.keys(dayCount).length > 0) {
    var maxDayCount = 0;
    var favoriteDay = 0;
    for (var day in dayCount) {
      if (dayCount[day] > maxDayCount) {
        maxDayCount = dayCount[day];
        favoriteDay = parseInt(day);
      }
    }
    var dayNames = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
    result.patterns.favoriteDay = dayNames[favoriteDay];
  }
  
  // === FIND FAVORITE TIME ===
  if (Object.keys(hourCount).length > 0) {
    var maxHourCount = 0;
    var favoriteHour = 0;
    for (var hour in hourCount) {
      if (hourCount[hour] > maxHourCount) {
        maxHourCount = hourCount[hour];
        favoriteHour = parseInt(hour);
      }
    }
    result.patterns.favoriteTime = favoriteHour + ":00 - " + (favoriteHour + 1) + ":00";
  }
  
  // === CALCULATE AVG DAYS BETWEEN VISITS ===
  if (result.visits.dates.length > 1) {
    var sortedDates = result.visits.dates.sort(function(a, b) { return new Date(a) - new Date(b); });
    var totalDays = 0;
    for (var i = 1; i < sortedDates.length; i++) {
      totalDays += (new Date(sortedDates[i]) - new Date(sortedDates[i-1])) / (1000 * 60 * 60 * 24);
    }
    result.patterns.avgDaysBetween = (totalDays / (sortedDates.length - 1)).toFixed(1) + " days";
  }
  
  // === BUILD TRANSACTION LIST ===
  for (var transId in transactionDetails) {
    var trans = transactionDetails[transId];
    var dateStr = Utilities.formatDate(new Date(trans.date), Session.getScriptTimeZone(), "MM/dd/yyyy");
    var items = trans.items.join(", ");
    var categories = Array.from(trans.categories).join(", ");
    var duration = bookingsByDate[Utilities.formatDate(new Date(trans.date), Session.getScriptTimeZone(), "yyyy-MM-dd")] || 0;
    
    result.transactions.push([
      dateStr,
      items,
      categories,
      trans.amount,
      duration > 0 ? duration.toFixed(1) + " hrs" : "",
      ""
    ]);
  }
  
  // Sort transactions by date (newest first)
  result.transactions.sort(function(a, b) {
    return new Date(b[0]) - new Date(a[0]);
  });
  
  return result;
}

/**
 * Generate insights based on member data
 */
function generateMemberInsights(data) {
  var insights = [];
  
  // Frequency insights
  if (data.visits.last30Days >= 8) {
    insights.push("🔥 HIGH FREQUENCY: Visits " + data.visits.last30Days + " times in last 30 days (2+ times/week)");
  } else if (data.visits.last30Days >= 4) {
    insights.push("✅ REGULAR VISITOR: Visits " + data.visits.last30Days + " times in last 30 days (weekly)");
  } else if (data.visits.last30Days <= 1 && data.visits.last30Days > 0) {
    insights.push("⚠️ LOW ACTIVITY: Only " + data.visits.last30Days + " visit(s) in last 30 days");
  } else if (data.visits.last30Days === 0) {
    insights.push("🚨 INACTIVE: No visits in last 30 days");
  }
  
  // Spending insights
  if (data.spending.avgPerVisit >= 100) {
    insights.push("💰 HIGH VALUE: Avg spend of $" + data.spending.avgPerVisit.toFixed(2) + " per visit");
  } else if (data.spending.avgPerVisit >= 50) {
    insights.push("💵 GOOD VALUE: Avg spend of $" + data.spending.avgPerVisit.toFixed(2) + " per visit");
  }
  
  // F&B insights
  if (data.fb.percentage >= 0.4) {
    insights.push("🍔 F&B ENTHUSIAST: " + (data.fb.percentage * 100).toFixed(0) + "% of spending is F&B");
  } else if (data.fb.percentage <= 0.1 && data.fb.total > 0) {
    insights.push("⛳ GOLF FOCUSED: Only " + (data.fb.percentage * 100).toFixed(0) + "% spending on F&B - upsell opportunity!");
  } else if (data.fb.total === 0 && data.visits.total > 0) {
    insights.push("🚨 NO F&B PURCHASES: Strong upsell opportunity!");
  }
  
  // Play time insights
  if (data.playTime.avgPerVisit >= 2) {
    insights.push("⏱️ LONG SESSIONS: Avg " + data.playTime.avgPerVisit.toFixed(1) + " hours per visit");
  } else if (data.playTime.avgPerVisit >= 1 && data.playTime.avgPerVisit < 2) {
    insights.push("⏱️ STANDARD SESSIONS: Avg " + data.playTime.avgPerVisit.toFixed(1) + " hours per visit");
  }
  
  // Pattern insights
  if (data.patterns.favoriteDay !== "N/A" && data.patterns.favoriteTime !== "N/A") {
    insights.push("📅 PREFERRED TIME: " + data.patterns.favoriteDay + "s at " + data.patterns.favoriteTime);
  }
  
  // Consistency insights
  if (data.patterns.avgDaysBetween !== "N/A") {
    var avgDays = parseFloat(data.patterns.avgDaysBetween);
    if (avgDays <= 7) {
      insights.push("🎯 VERY CONSISTENT: Returns every " + data.patterns.avgDaysBetween + " on average");
    } else if (avgDays <= 14) {
      insights.push("✅ CONSISTENT: Returns every " + data.patterns.avgDaysBetween + " on average");
    }
  }
  
  // Membership insights
  if (data.categories.membership > 0) {
    insights.push("👤 CURRENT MEMBER: Has purchased membership");
  } else if (data.visits.total >= 4) {
    insights.push("💎 NOT A MEMBER: Excellent candidate for membership! (visits " + data.visits.total + " times)");
  } else {
    insights.push("💎 NOT A MEMBER: Could be membership candidate with more visits");
  }
  
  return insights;
}

// ============================================
// FUNCTION 2: FIND MEMBERSHIP LEADS
// ============================================

/**
 * Main function - find and rank non-members by potential
 */
function findMembershipLeads() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var ui = SpreadsheetApp.getUi();
  
  ui.alert(
    '🔍 Finding Membership Leads...',
    'This will analyze all non-member customers and rank them by membership potential.\n\nThis may take a minute...',
    ui.ButtonSet.OK
  );
  
  var leads = analyzeMembershipLeads();
  
  displayMembershipLeads(leads);
}

/**
 * Analyze all customers and score them for membership potential
 */
function analyzeMembershipLeads() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var now = new Date();
  var ninetyDaysAgo = new Date(now.getTime() - (90 * 24 * 60 * 60 * 1000));
  
  // === GET CURRENT MEMBERS ===
  var memberSet = getCurrentMembers();
  
  // === GET ALL CUSTOMERS ===
  var customerSheet = ss.getSheetByName("Square Customer Export");
  var customerData = customerSheet.getDataRange().getValues();
  var customerHeaders = customerData[0];
  
  var customerIdCol = customerHeaders.indexOf("Square Customer ID");
  var firstNameCol = customerHeaders.indexOf("First Name");
  var lastNameCol = customerHeaders.indexOf("Last Name");
  var emailCol = customerHeaders.indexOf("Email Address");
  
  // === GET TRANSACTIONS ===
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];

  var transCustomerIdCol = transHeaders.indexOf("Customer ID");
  var transEmailCol = transHeaders.indexOf("Customer Email");
  var transDateCol = transHeaders.indexOf("Date");
  var transCollectedCol = transHeaders.indexOf("Total Collected");
  var transIdCol = transHeaders.indexOf("Transaction ID");
  
  // === GET ITEMS FOR F&B TRACKING ===
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemNameCol = itemHeaders.indexOf("Item");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  var itemGrossCol = itemHeaders.indexOf("Gross Sales");
  
  // Build F&B lookup by transaction
  var fbByTransaction = {};
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    var itemName = itemData[i][itemNameCol];
    var category = itemData[i][itemCategoryCol];
    var grossSales = parseFloat(itemData[i][itemGrossCol]) || 0;
    
    var majorCat = getMajorCategory(category, itemName);
    
    if (majorCat === "Food" || majorCat === "Beverage") {
      fbByTransaction[transId] = (fbByTransaction[transId] || 0) + grossSales;
    }
  }
  
  // === GET BOOKINGS DATA FOR VISIT TRACKING ===
  var bookingSheet = ss.getSheetByName("Apex Bookings Export");
  var bookingsByCustomer = {}; // Map email -> array of booking dates

  if (bookingSheet) {
    var bookingData = bookingSheet.getDataRange().getValues();
    var bookingHeaders = bookingData[0];

    var bookingEmailCol = bookingHeaders.indexOf("Email");
    var bookingDateCol = bookingHeaders.indexOf("Date");

    for (var i = 1; i < bookingData.length; i++) {
      var bookingEmail = normalizeEmail(bookingData[i][bookingEmailCol]);
      var bookingDate = new Date(bookingData[i][bookingDateCol]);

      if (bookingEmail && bookingDate >= ninetyDaysAgo) {
        if (!bookingsByCustomer[bookingEmail]) {
          bookingsByCustomer[bookingEmail] = [];
        }
        bookingsByCustomer[bookingEmail].push(bookingDate);
      }
    }
  }

  // === ANALYZE EACH NON-MEMBER ===
  var leads = [];

  for (var i = 1; i < customerData.length; i++) {
    var customerId = customerData[i][customerIdCol];

    // Skip if already a member
    if (memberSet.has(customerId)) {
      continue;
    }

    var firstName = customerData[i][firstNameCol] || "";
    var lastName = customerData[i][lastNameCol] || "";
    var email = customerData[i][emailCol] || "";
    var fullName = (firstName + " " + lastName).trim() || email;

    // Calculate metrics for last 90 days
    var visits90 = 0;
    var spending90 = 0;
    var fbSpending90 = 0;
    var visitDates = [];
    var visitDateSet90 = {}; // Track unique visit dates in last 90 days (bookings + transactions)
    var normEmail = normalizeEmail(email);

    // First, add booking visits for this customer
    if (normEmail && bookingsByCustomer[normEmail]) {
      for (var k = 0; k < bookingsByCustomer[normEmail].length; k++) {
        var bookingDate = bookingsByCustomer[normEmail][k];
        var dateStr = Utilities.formatDate(bookingDate, Session.getScriptTimeZone(), "yyyy-MM-dd");

        if (!visitDateSet90[dateStr]) {
          visitDateSet90[dateStr] = bookingDate;
          visitDates.push(bookingDate);
        }
      }
    }

    // Then add transaction visits
    for (var j = 1; j < transData.length; j++) {
      var transCustomerId = transData[j][transCustomerIdCol];
      var transEmail = normalizeEmail(transData[j][transEmailCol]);

      // Match by BOTH Customer ID AND email to catch all transactions
      if (transCustomerId === customerId || (normEmail && transEmail === normEmail)) {
        var date = new Date(transData[j][transDateCol]);

        if (date >= ninetyDaysAgo) {
          var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");

          // Track unique visit dates
          if (!visitDateSet90[dateStr]) {
            visitDateSet90[dateStr] = date;
            visitDates.push(date);
          }

          var amount = parseFloat(transData[j][transCollectedCol]) || 0;
          spending90 += amount;

          var transId = transData[j][transIdCol];
          if (fbByTransaction[transId]) {
            fbSpending90 += fbByTransaction[transId];
          }
        }
      }
    }

    // Count unique visits (from both bookings and transactions)
    visits90 = Object.keys(visitDateSet90).length;
    
    // Skip if no recent activity
    if (visits90 === 0) {
      continue;
    }
    
    var avgSpendPerVisit = spending90 / visits90;
    var visitsPerMonth = (visits90 / 90) * 30;
    var fbPercent = spending90 > 0 ? (fbSpending90 / spending90) : 0;
    
    // === CALCULATE SCORES ===
    
    // Frequency score (0-100): 10+ visits/month = 100
    var frequencyScore = Math.min((visitsPerMonth / 10) * 100, 100);
    
    // Spending score (0-100): $100/visit = 100
    var spendingScore = Math.min((avgSpendPerVisit / 100) * 100, 100);
    
    // F&B score (0-100): Higher F&B engagement = better member
    var fbScore = fbPercent * 100;
    
    // Consistency score (0-100): Regular intervals = higher score
    var consistencyScore = 0;
    if (visitDates.length > 1) {
      var sortedDates = visitDates.sort(function(a, b) { return a - b; });
      var intervals = [];
      for (var k = 1; k < sortedDates.length; k++) {
        intervals.push((sortedDates[k] - sortedDates[k-1]) / (1000 * 60 * 60 * 24));
      }
      
      // Calculate standard deviation of intervals
      var avgInterval = intervals.reduce(function(a, b) { return a + b; }, 0) / intervals.length;
      var variance = 0;
      for (var k = 0; k < intervals.length; k++) {
        variance += Math.pow(intervals[k] - avgInterval, 2);
      }
      variance = variance / intervals.length;
      var stdDev = Math.sqrt(variance);
      
      // Lower std dev = more consistent = higher score
      consistencyScore = Math.max(0, 100 - (stdDev * 2));
    }
    
    // === OVERALL SCORE (weighted average) ===
    var overallScore = (
      frequencyScore * 0.4 +
      spendingScore * 0.3 +
      fbScore * 0.15 +
      consistencyScore * 0.15
    );
    
    leads.push({
      customerId: customerId,
      name: fullName,
      email: email,
      visits90: visits90,
      visitsPerMonth: visitsPerMonth.toFixed(1),
      spending90: spending90,
      avgSpend: avgSpendPerVisit,
      fbSpending: fbSpending90,
      fbPercent: fbPercent,
      frequencyScore: frequencyScore,
      spendingScore: spendingScore,
      fbScore: fbScore,
      consistencyScore: consistencyScore,
      overallScore: overallScore
    });
  }
  
  // Sort by overall score
  leads.sort(function(a, b) {
    return b.overallScore - a.overallScore;
  });
  
  return leads;
}

/**
 * Get set of current member IDs
 * Checks for Putter, Iron, or Driver membership specifically
 */
function getCurrentMembers() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var memberSet = new Set();
  
  // Find all customers who have purchased Putter, Iron, or Driver memberships
  var itemSheet = ss.getSheetByName("Square Item Detail Export");
  var itemData = itemSheet.getDataRange().getValues();
  var itemHeaders = itemData[0];
  
  var itemTransIdCol = itemHeaders.indexOf("Transaction ID");
  var itemNameCol = itemHeaders.indexOf("Item");
  var itemCategoryCol = itemHeaders.indexOf("Category");
  
  var transSheet = ss.getSheetByName("Square Transactions Export");
  var transData = transSheet.getDataRange().getValues();
  var transHeaders = transData[0];
  
  var transIdCol = transHeaders.indexOf("Transaction ID");
  var transCustomerIdCol = transHeaders.indexOf("Customer ID");
  
  // Build transaction to customer map
  var transToCustomer = {};
  for (var i = 1; i < transData.length; i++) {
    var transId = transData[i][transIdCol];
    var customerId = transData[i][transCustomerIdCol];
    if (transId && customerId) {
      transToCustomer[transId] = customerId;
    }
  }
  
  // Find membership purchases - specifically looking for Putter, Iron, or Driver
  for (var i = 1; i < itemData.length; i++) {
    var transId = itemData[i][itemTransIdCol];
    var itemName = String(itemData[i][itemNameCol] || "").toLowerCase();
    var category = itemData[i][itemCategoryCol];
    
    var majorCat = getMajorCategory(category, itemData[i][itemNameCol]);
    
    // Check if this is a Putter, Iron, or Driver membership
    var isPutterMember = itemName.indexOf("putter") >= 0 && (itemName.indexOf("membership") >= 0 || itemName.indexOf("member") >= 0);
    var isIronMember = itemName.indexOf("iron") >= 0 && (itemName.indexOf("membership") >= 0 || itemName.indexOf("member") >= 0);
    var isDriverMember = itemName.indexOf("driver") >= 0 && (itemName.indexOf("membership") >= 0 || itemName.indexOf("member") >= 0);
    
    // Also check if it's categorized as Membership
    var isMembershipCategory = majorCat === "Membership";
    
    if (isPutterMember || isIronMember || isDriverMember || isMembershipCategory) {
      var customerId = transToCustomer[transId];
      if (customerId) {
        memberSet.add(customerId);
      }
    }
  }
  
  return memberSet;
}

/**
 * Display membership leads in a formatted sheet
 */
function displayMembershipLeads(leads) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var leadsSheet = ss.getSheetByName("Membership Leads");
  if (leadsSheet) {
    leadsSheet.clear();
  } else {
    leadsSheet = ss.insertSheet("Membership Leads");
  }
  
  // === HEADER ===
  leadsSheet.getRange("A1:L1").merge();
  leadsSheet.getRange("A1").setValue("💎 MEMBERSHIP LEAD RECOMMENDATIONS");
  leadsSheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
  leadsSheet.getRange("A1").setBackground("#9C27B0").setFontColor("white");
  
  leadsSheet.getRange("A2:L2").merge();
  leadsSheet.getRange("A2").setValue("Ranked by membership potential | Last 90 Days Data | Generated: " + new Date());
  leadsSheet.getRange("A2").setFontSize(10).setHorizontalAlignment("center").setBackground("#f3e5f5");
  
  // === COLUMN HEADERS ===
  var headers = [
    "Rank",
    "Name",
    "Email",
    "Visits (90d)",
    "Visits/Month",
    "Total Spend",
    "Avg/Visit",
    "F&B Spend",
    "F&B %",
    "Overall Score",
    "Priority",
    "Notes"
  ];
  
  leadsSheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  leadsSheet.getRange(4, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8E8E8");
  leadsSheet.setFrozenRows(4);
  
  // === DATA ===
  var outputData = [];
  var hotCount = 0;
  var goodCount = 0;
  var warmCount = 0;
  var coldCount = 0;
  
  for (var i = 0; i < leads.length; i++) {
    var lead = leads[i];
    
    var priority;
    var notes = [];
    
    if (lead.overallScore >= 70) {
      priority = "🔥 HOT";
      notes.push("Strong candidate");
      hotCount++;
    } else if (lead.overallScore >= 50) {
      priority = "✅ GOOD";
      notes.push("Good potential");
      goodCount++;
    } else if (lead.overallScore >= 30) {
      priority = "⚠️ WARM";
      notes.push("Needs nurturing");
      warmCount++;
    } else {
      priority = "❄️ COLD";
      notes.push("Low priority");
      coldCount++;
    }
    
    if (lead.visitsPerMonth >= 8) notes.push("High frequency");
    if (lead.avgSpend >= 80) notes.push("High spender");
    if (lead.fbPercent >= 0.3) notes.push("F&B engaged");
    if (lead.consistencyScore >= 70) notes.push("Consistent");
    
    outputData.push([
      i + 1,
      lead.name,
      lead.email,
      lead.visits90,
      lead.visitsPerMonth,
      lead.spending90,
      lead.avgSpend,
      lead.fbSpending,
      lead.fbPercent,
      lead.overallScore.toFixed(0),
      priority,
      notes.join(", ")
    ]);
  }
  
  if (outputData.length > 0) {
    leadsSheet.getRange(5, 1, outputData.length, headers.length).setValues(outputData);
    
    // Format currency and percentage columns
    leadsSheet.getRange(5, 6, outputData.length, 1).setNumberFormat("$#,##0.00");
    leadsSheet.getRange(5, 7, outputData.length, 1).setNumberFormat("$#,##0.00");
    leadsSheet.getRange(5, 8, outputData.length, 1).setNumberFormat("$#,##0.00");
    leadsSheet.getRange(5, 9, outputData.length, 1).setNumberFormat("0.0%");
    
    // Color code by priority
    for (var i = 0; i < outputData.length; i++) {
      var row = i + 5;
      var priority = outputData[i][10];
      var color;
      
      if (priority.indexOf("HOT") >= 0) {
        color = "#f4cccc";
      } else if (priority.indexOf("GOOD") >= 0) {
        color = "#d9ead3";
      } else if (priority.indexOf("WARM") >= 0) {
        color = "#fff2cc";
      } else {
        color = "#ffffff";
      }
      
      leadsSheet.getRange(row, 1, 1, headers.length).setBackground(color);
    }
    
    // Add borders
    leadsSheet.getRange(4, 1, outputData.length + 1, headers.length).setBorder(true, true, true, true, true, true);
    
    // Highlight top 3
    if (outputData.length >= 1) leadsSheet.getRange(5, 1, 1, headers.length).setBackground("#ea9999").setFontWeight("bold");
    if (outputData.length >= 2) leadsSheet.getRange(6, 1, 1, headers.length).setBackground("#f9cb9c").setFontWeight("bold");
    if (outputData.length >= 3) leadsSheet.getRange(7, 1, 1, headers.length).setBackground("#ffe599").setFontWeight("bold");
  }
  
  // Auto-resize columns
  for (var i = 1; i <= headers.length; i++) {
    leadsSheet.autoResizeColumn(i);
  }
  
  ss.setActiveSheet(leadsSheet);
  
  // === SUMMARY MESSAGE ===
  var summary = '✅ Lead Analysis Complete!\n\n';
  summary += 'Found ' + leads.length + ' potential members\n\n';
  summary += '🔥 HOT leads: ' + hotCount + '\n';
  summary += '✅ GOOD leads: ' + goodCount + '\n';
  summary += '⚠️ WARM leads: ' + warmCount + '\n';
  summary += '❄️ COLD leads: ' + coldCount + '\n\n';
  
  if (leads.length > 0) {
    summary += 'Top 3 Leads:\n';
    if (leads[0]) summary += '1. ' + leads[0].name + ' (Score: ' + leads[0].overallScore.toFixed(0) + ')\n';
    if (leads[1]) summary += '2. ' + leads[1].name + ' (Score: ' + leads[1].overallScore.toFixed(0) + ')\n';
    if (leads[2]) summary += '3. ' + leads[2].name + ' (Score: ' + leads[2].overallScore.toFixed(0) + ')\n';
  }
  
  summary += '\nCheck "Membership Leads" sheet for full ranking!';
  
  SpreadsheetApp.getUi().alert('Membership Leads', summary, SpreadsheetApp.getUi().ButtonSet.OK);
}