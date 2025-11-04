/**
 * ========================================
 * APEX GOLF ANALYTICS - FILE 6: TIP CALCULATOR
 * ========================================
 * 
 * THIS FILE CONTAINS:
 * - Tip distribution calculator based on employee work overlap
 * - Payroll summary for tip entry
 * 
 * Add to menu in File 1 onOpen():
 * .addItem('💰 Calculate Tip Distribution', 'calculateTipDistribution')
 */

// ============================================
// MAIN FUNCTION: CALCULATE TIP DISTRIBUTION
// ============================================

/**
 * Main function to calculate tip distribution based on employee overlap with customer bookings
 */
function calculateTipDistribution() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var ui = SpreadsheetApp.getUi();
    
    ui.alert(
      '💰 Calculating Tip Distribution...',
      'This will analyze all tips and distribute them based on employee work overlap.\n\nThis may take a minute...',
      ui.ButtonSet.OK
    );
    
    // Get all the data we need
    Logger.log("Getting staff shifts...");
    var staffShifts = getStaffShifts();
    Logger.log("Found " + staffShifts.length + " staff shifts");
    
    // Determine pay period from timecards
    var payPeriod = getPayPeriodFromShifts(staffShifts);
    Logger.log("Pay period: " + formatDateTime(payPeriod.start) + " to " + formatDateTime(payPeriod.end));
    
    Logger.log("Getting transactions with tips...");
    var transactions = getTransactionsWithTips(payPeriod);
    Logger.log("Found " + transactions.length + " transactions with tips in pay period");
    
    Logger.log("Getting bookings...");
    var bookings = getBookings(payPeriod);
    Logger.log("Found " + bookings.length + " bookings in pay period");
    
    // Calculate tip distribution
    Logger.log("Calculating distribution...");
    var tipDistribution = calculateDistribution(staffShifts, transactions, bookings, payPeriod);
    
    // Display results
    Logger.log("Displaying results...");
    displayTipResults(tipDistribution, payPeriod);
    
  } catch (error) {
    Logger.log("ERROR: " + error.toString());
    SpreadsheetApp.getUi().alert(
      '❌ Error',
      'An error occurred:\n\n' + error.toString() + '\n\nCheck View > Logs for details.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}

// ============================================
// DATA GATHERING FUNCTIONS
// ============================================

/**
 * Get all staff shifts from Staff Timecards
 */
function getStaffShifts() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var staffSheet = ss.getSheetByName("Staff Timecards");
  
  if (!staffSheet) {
    throw new Error('Cannot find "Staff Timecards" sheet! Please make sure it exists.');
  }
  
  var lastRow = staffSheet.getLastRow();
  var lastCol = staffSheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error('Staff Timecards sheet appears to be empty!');
  }
  
  Logger.log("Staff Timecards: " + lastRow + " rows, " + lastCol + " columns");
  
  var data = staffSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  
  var firstNameCol = headers.indexOf("First name");
  var lastNameCol = headers.indexOf("Last name");
  var jobTitleCol = headers.indexOf("Job title");
  var clockInDateCol = headers.indexOf("Clockin date");
  var clockInTimeCol = headers.indexOf("Clockin time");
  var clockOutDateCol = headers.indexOf("Clockout date");
  var clockOutTimeCol = headers.indexOf("Clockout time");
  
  if (firstNameCol === -1 || lastNameCol === -1) {
    throw new Error('Cannot find required columns in Staff Timecards sheet! Looking for: First name, Last name, Clockin date, Clockin time, Clockout date, Clockout time');
  }
  
  var shifts = [];
  
  for (var i = 1; i < data.length; i++) {
    var firstName = data[i][firstNameCol] || "";
    var lastName = data[i][lastNameCol] || "";
    var fullName = (firstName + " " + lastName).trim();
    
    if (!fullName) continue;
    
    var clockInDate = data[i][clockInDateCol];
    var clockInTime = data[i][clockInTimeCol];
    var clockOutDate = data[i][clockOutDateCol];
    var clockOutTime = data[i][clockOutTimeCol];
    
    if (!clockInDate || !clockOutDate) continue;
    
    // Parse clock in/out times
    var clockIn = parseDateTimeFromSquare(clockInDate, clockInTime);
    var clockOut = parseDateTimeFromSquare(clockOutDate, clockOutTime);
    
    if (!clockIn || !clockOut) continue;
    
    shifts.push({
      firstName: firstName,
      lastName: lastName,
      fullName: fullName,
      jobTitle: data[i][jobTitleCol] || "",
      clockIn: clockIn,
      clockOut: clockOut,
      isTipEligible: fullName.toLowerCase() !== "bronson roberts"
    });
  }
  
  Logger.log("Parsed " + shifts.length + " valid shifts");
  return shifts;
}

/**
 * Determine pay period start and end from staff shifts
 */
function getPayPeriodFromShifts(shifts) {
  if (shifts.length === 0) {
    throw new Error('No shifts found - cannot determine pay period');
  }
  
  var earliestClockIn = shifts[0].clockIn;
  var latestClockOut = shifts[0].clockOut;
  
  for (var i = 1; i < shifts.length; i++) {
    if (shifts[i].clockIn < earliestClockIn) {
      earliestClockIn = shifts[i].clockIn;
    }
    if (shifts[i].clockOut > latestClockOut) {
      latestClockOut = shifts[i].clockOut;
    }
  }
  
  // Set to start of day for earliest and end of day for latest
  var startDate = new Date(earliestClockIn);
  startDate.setHours(0, 0, 0, 0);
  
  var endDate = new Date(latestClockOut);
  endDate.setHours(23, 59, 59, 999);
  
  return {
    start: startDate,
    end: endDate
  };
}

/**
 * Get all transactions with tips (filtered by pay period)
 */
function getTransactionsWithTips(payPeriod) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var transSheet = ss.getSheetByName("Square Transactions Export");
  
  if (!transSheet) {
    throw new Error('Cannot find "Square Transactions Export" sheet! Please make sure it exists.');
  }
  
  var lastRow = transSheet.getLastRow();
  var lastCol = transSheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error('Square Transactions Export sheet appears to be empty!');
  }
  
  Logger.log("Square Transactions Export: " + lastRow + " rows, " + lastCol + " columns");
  
  var data = transSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  
  var transIdCol = headers.indexOf("Transaction ID");
  var dateCol = headers.indexOf("Date");
  var timeCol = headers.indexOf("Time");
  var tipCol = headers.indexOf("Tip");
  var customerIdCol = headers.indexOf("Customer ID");
  var customerNameCol = headers.indexOf("Customer Name");
  var customerEmailCol = headers.indexOf("Customer Email");
  
  if (transIdCol === -1 || dateCol === -1 || tipCol === -1) {
    throw new Error('Cannot find required columns in Square Transactions Export! Looking for: Transaction ID, Date, Tip');
  }
  
  var transactions = [];
  
  for (var i = 1; i < data.length; i++) {
    var tip = parseFloat(data[i][tipCol]) || 0;
    
    // Include ALL tips - positive, negative, and zero
    // Negative tips are refunds/voids that need to be accounted for
    if (tip === 0) continue; // Skip only zero tips (no tip given)
    
    var date = data[i][dateCol];
    var time = data[i][timeCol];
    
    if (!date) continue;
    
    var transDateTime = combineDateAndTime(date, time);
    
    // Filter by pay period
    if (transDateTime < payPeriod.start || transDateTime > payPeriod.end) {
      continue;
    }
    
    transactions.push({
      transactionId: data[i][transIdCol],
      date: date,
      time: time,
      dateTime: transDateTime,
      tip: tip,
      customerId: data[i][customerIdCol] || "",
      customerName: data[i][customerNameCol] || "",
      customerEmail: normalizeEmail(data[i][customerEmailCol])
    });
  }
  
  Logger.log("Found " + transactions.length + " transactions with tips in pay period (including refunds/voids)");
  
  // Log first 3 transactions as samples
  if (transactions.length > 0) {
    Logger.log("Sample transactions (first 3):");
    for (var i = 0; i < Math.min(3, transactions.length); i++) {
      var t = transactions[i];
      Logger.log("  " + (i+1) + ". ID: " + t.transactionId + " | Email: '" + t.customerEmail + "' | Name: '" + t.customerName + "' | Tip: $" + t.tip.toFixed(2) + " | " + formatDateTime(t.dateTime));
    }
    
    // Count negative tips
    var negativeCount = 0;
    var negativeTotal = 0;
    for (var i = 0; i < transactions.length; i++) {
      if (transactions[i].tip < 0) {
        negativeCount++;
        negativeTotal += transactions[i].tip;
      }
    }
    if (negativeCount > 0) {
      Logger.log("⚠️ Found " + negativeCount + " negative tips (refunds/voids): $" + negativeTotal.toFixed(2));
    }
  }
  
  return transactions;
}

/**
 * Get all bookings (filtered by pay period)
 */
function getBookings(payPeriod) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var bookingSheet = ss.getSheetByName("Apex Bookings Export");
  
  if (!bookingSheet) {
    throw new Error('Cannot find "Apex Bookings Export" sheet! Please make sure it exists.');
  }
  
  var lastRow = bookingSheet.getLastRow();
  var lastCol = bookingSheet.getLastColumn();
  
  if (lastRow < 2) {
    throw new Error('Apex Bookings Export sheet appears to be empty!');
  }
  
  Logger.log("Apex Bookings Export: " + lastRow + " rows, " + lastCol + " columns");
  
  var data = bookingSheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = data[0];
  
  var emailCol = headers.indexOf("Email");
  var dateCol = headers.indexOf("Date");
  var timeCol = headers.indexOf("Time");
  var durationCol = headers.indexOf("Duration Mins");
  var firstNameCol = headers.indexOf("First Name");
  var lastNameCol = headers.indexOf("Last Name");
  
  if (emailCol === -1 || dateCol === -1 || durationCol === -1) {
    throw new Error('Cannot find required columns in Apex Bookings Export! Looking for: Email, Date, Duration Mins');
  }
  
  var bookings = [];
  var fixedDurationCount = 0;
  
  for (var i = 1; i < data.length; i++) {
    var email = normalizeEmail(data[i][emailCol]);
    var date = data[i][dateCol];
    var time = data[i][timeCol];
    var duration = parseFloat(data[i][durationCol]) || 0;
    
    if (!date) continue;
    
    // If duration is 0 or invalid, use default 60 minutes
    if (duration <= 0) {
      duration = 60; // Default to 1 hour
      fixedDurationCount++;
    }
    
    var startTime = combineDateAndTime(date, time);
    if (!startTime) continue;
    
    // Filter by pay period
    if (startTime < payPeriod.start || startTime > payPeriod.end) {
      continue;
    }
    
    var endTime = new Date(startTime.getTime() + (duration * 60 * 1000));
    
    bookings.push({
      email: email,
      firstName: data[i][firstNameCol] || "",
      lastName: data[i][lastNameCol] || "",
      date: date,
      startTime: startTime,
      endTime: endTime,
      durationMins: duration
    });
  }
  
  if (fixedDurationCount > 0) {
    Logger.log("Fixed " + fixedDurationCount + " bookings with 0 or invalid duration (assigned 60 min default)");
  }
  
  Logger.log("Found " + bookings.length + " valid bookings in pay period");
  
  // Log first 3 bookings as samples
  if (bookings.length > 0) {
    Logger.log("Sample bookings (first 3):");
    for (var i = 0; i < Math.min(3, bookings.length); i++) {
      var b = bookings[i];
      Logger.log("  " + (i+1) + ". Email: " + b.email + " | " + formatDateTime(b.startTime) + " to " + formatDateTime(b.endTime) + " (" + b.durationMins + " mins)");
    }
  }
  
  return bookings;
}

// ============================================
// CALCULATION FUNCTIONS
// ============================================

/**
 * Calculate tip distribution for all transactions
 */
function calculateDistribution(staffShifts, transactions, bookings, payPeriod) {
  var results = {
    details: [],
    summary: {},
    overpaid: 0,
    totalTipsDistributed: 0,
    totalTipsProcessed: 0
  };
  
  Logger.log("Starting distribution calculation for " + transactions.length + " transactions");
  Logger.log("Using proximity-based booking matching...");
  
  var transactionCount = 0;
  var bookingFoundCount = 0;
  var noBookingCount = 0;
  var processedTransactionIds = {}; // Track processed IDs to detect duplicates
  var duplicateCount = 0;
  
  // Process each transaction with tips
  for (var i = 0; i < transactions.length; i++) {
    var trans = transactions[i];
    
    // Check for duplicate transaction IDs
    if (processedTransactionIds[trans.transactionId]) {
      Logger.log("⚠️ DUPLICATE TRANSACTION: " + trans.transactionId + " - SKIPPING");
      duplicateCount++;
      continue;
    }
    processedTransactionIds[trans.transactionId] = true;
    
    transactionCount++;
    
    // Find the closest booking(s) to this transaction time
    var closestBooking = findClosestBooking(trans, bookings);
    
    if (!closestBooking) {
      noBookingCount++;
      // No booking found - split evenly among clocked-in employees at payment time
      Logger.log("No booking near transaction " + trans.transactionId + " - splitting among clocked-in employees");
      
      var paymentTime = trans.dateTime;
      var clockedInEmployees = findEmployeesAtTime(staffShifts, paymentTime);
      
      if (clockedInEmployees.length === 0) {
        results.details.push({
          transactionId: trans.transactionId,
          customerName: trans.customerName,
          transactionDateTime: trans.dateTime,
          tip: trans.tip,
          bookingStart: paymentTime,
          bookingEnd: paymentTime,
          employeeDistribution: [],
          overpaid: trans.tip,
          reason: "No booking found - No employees clocked in at payment time"
        });
        results.overpaid += trans.tip;
        results.totalTipsProcessed += trans.tip;
        continue;
      }
      
      // Check if only Bronson was clocked in
      var tipEligibleEmployees = clockedInEmployees.filter(function(emp) { 
        return emp.isTipEligible; 
      });
      
      if (tipEligibleEmployees.length === 0) {
        Logger.log("⚠️ Only Bronson worked - Transaction: " + trans.transactionId + " | Tip: $" + trans.tip.toFixed(2) + " → Overpaid");
        
        results.details.push({
          transactionId: trans.transactionId,
          customerName: trans.customerName,
          transactionDateTime: trans.dateTime,
          tip: trans.tip,
          bookingStart: paymentTime,
          bookingEnd: paymentTime,
          employeeDistribution: clockedInEmployees.map(function(emp) {
            return {
              employeeName: emp.fullName,
              clockIn: emp.clockIn,
              clockOut: emp.clockOut,
              overlapPercent: 100,
              tipAmount: 0,
              note: "Not tip-eligible (Bronson Roberts)"
            };
          }),
          overpaid: trans.tip,
          reason: "No booking - Only Bronson Roberts clocked in (not tip-eligible)"
        });
        results.overpaid += trans.tip;
        results.totalTipsProcessed += trans.tip;
        continue;
      }
      
      // Split evenly
      var tipPerEmployee = trans.tip / tipEligibleEmployees.length;
      var employeeDistribution = [];
      
      for (var k = 0; k < clockedInEmployees.length; k++) {
        var emp = clockedInEmployees[k];
        var tipAmount = 0;
        var note = "";
        
        if (emp.isTipEligible) {
          tipAmount = tipPerEmployee;
          if (!results.summary[emp.fullName]) {
            results.summary[emp.fullName] = 0;
          }
          results.summary[emp.fullName] += tipAmount;
          results.totalTipsDistributed += tipAmount;
        } else {
          note = "Not tip-eligible (Bronson Roberts)";
        }
        
        employeeDistribution.push({
          employeeName: emp.fullName,
          clockIn: emp.clockIn,
          clockOut: emp.clockOut,
          overlapPercent: 100,
          tipAmount: tipAmount,
          note: note
        });
      }
      
      results.details.push({
        transactionId: trans.transactionId,
        customerName: trans.customerName,
        transactionDateTime: trans.dateTime,
        tip: trans.tip,
        bookingStart: paymentTime,
        bookingEnd: paymentTime,
        employeeDistribution: employeeDistribution,
        overpaid: 0,
        reason: "⚠️ No booking found - Split evenly among " + tipEligibleEmployees.length + " clocked-in employee(s)"
      });
      
      results.totalTipsProcessed += trans.tip;
      continue;
    }
    
    // BOOKING FOUND - Use pro-rated distribution based on overlap
    bookingFoundCount++;
    
    // IMPORTANT: Use transaction time as the END of the work period
    // Customers pay when they're done, so only count employees who worked up to payment time
    var workStartTime = closestBooking.startTime;
    var workEndTime = trans.dateTime; // Use transaction time, not booking end time!
    
    var workingEmployees = findWorkingEmployees(staffShifts, workStartTime, workEndTime);
    
    if (workingEmployees.length === 0) {
      results.details.push({
        transactionId: trans.transactionId,
        customerName: trans.customerName,
        transactionDateTime: trans.dateTime,
        tip: trans.tip,
        bookingStart: workStartTime,
        bookingEnd: workEndTime,
        employeeDistribution: [],
        overpaid: trans.tip,
        reason: "No employees working from booking start to payment time"
      });
      results.overpaid += trans.tip;
      results.totalTipsProcessed += trans.tip;
      continue;
    }
    
    // Check if only Bronson worked
    var tipEligibleEmployees = workingEmployees.filter(function(emp) { 
      return emp.isTipEligible; 
    });
    
    if (tipEligibleEmployees.length === 0) {
      Logger.log("⚠️ Only Bronson worked - Transaction: " + trans.transactionId + " | Tip: $" + trans.tip.toFixed(2) + " | Work period: " + formatDateTime(workStartTime) + " to " + formatDateTime(workEndTime) + " → Overpaid");
      
      results.details.push({
        transactionId: trans.transactionId,
        customerName: trans.customerName,
        transactionDateTime: trans.dateTime,
        tip: trans.tip,
        bookingStart: workStartTime,
        bookingEnd: workEndTime, // Transaction time, not booking end time
        employeeDistribution: workingEmployees.map(function(emp) {
          return {
            employeeName: emp.fullName,
            clockIn: emp.clockIn,
            clockOut: emp.clockOut,
            overlapPercent: emp.overlapPercent,
            tipAmount: 0,
            note: "Not tip-eligible (Bronson Roberts)"
          };
        }),
        overpaid: trans.tip,
        reason: "Only Bronson Roberts worked from booking start to payment time - not tip-eligible"
      });
      results.overpaid += trans.tip;
      results.totalTipsProcessed += trans.tip;
      continue;
    }
    
    // Calculate total overlap percentage for tip-eligible employees only
    var totalEligibleOverlap = 0;
    for (var k = 0; k < tipEligibleEmployees.length; k++) {
      totalEligibleOverlap += tipEligibleEmployees[k].overlapPercent;
    }
    
    // Distribute tips proportionally
    var employeeDistribution = [];
    
    for (var k = 0; k < workingEmployees.length; k++) {
      var emp = workingEmployees[k];
      var tipAmount = 0;
      var note = "";
      
      if (emp.isTipEligible) {
        var proportion = emp.overlapPercent / totalEligibleOverlap;
        tipAmount = trans.tip * proportion;
        
        if (!results.summary[emp.fullName]) {
          results.summary[emp.fullName] = 0;
        }
        results.summary[emp.fullName] += tipAmount;
        results.totalTipsDistributed += tipAmount;
      } else {
        note = "Not tip-eligible (Bronson Roberts)";
      }
      
      employeeDistribution.push({
        employeeName: emp.fullName,
        clockIn: emp.clockIn,
        clockOut: emp.clockOut,
        overlapPercent: emp.overlapPercent,
        tipAmount: tipAmount,
        note: note
      });
    }
    
    results.details.push({
      transactionId: trans.transactionId,
      customerName: trans.customerName,
      transactionDateTime: trans.dateTime,
      tip: trans.tip,
      bookingStart: workStartTime,
      bookingEnd: workEndTime, // Transaction time
      employeeDistribution: employeeDistribution,
      overpaid: 0,
      reason: ""
    });
    
    results.totalTipsProcessed += trans.tip;
  }
  
  Logger.log("=== DISTRIBUTION SUMMARY ===");
  Logger.log("Total transactions in data: " + transactions.length);
  Logger.log("Duplicate transactions skipped: " + duplicateCount);
  Logger.log("Unique transactions processed: " + transactionCount);
  Logger.log("Bookings found: " + bookingFoundCount);
  Logger.log("No booking found: " + noBookingCount);
  Logger.log("Tips Processed: $" + results.totalTipsProcessed.toFixed(2));
  Logger.log("Tips Distributed: $" + results.totalTipsDistributed.toFixed(2));
  Logger.log("Tips Overpaid: $" + results.overpaid.toFixed(2));
  
  if (results.overpaid > 0) {
    Logger.log("⚠️ Overpaid tips breakdown:");
    var bronsonOnlyCount = 0;
    var noEmployeesCount = 0;
    for (var i = 0; i < results.details.length; i++) {
      if (results.details[i].overpaid > 0) {
        if (results.details[i].reason.indexOf("Bronson") >= 0) {
          bronsonOnlyCount++;
        } else if (results.details[i].reason.indexOf("No employees") >= 0) {
          noEmployeesCount++;
        }
      }
    }
    Logger.log("  - Only Bronson worked: " + bronsonOnlyCount + " transactions");
    Logger.log("  - No employees working: " + noEmployeesCount + " transactions");
  }
  
  // Validation check - tips distributed + overpaid should equal tips processed
  var calculatedTotal = results.totalTipsDistributed + results.overpaid;
  var difference = Math.abs(calculatedTotal - results.totalTipsProcessed);
  
  Logger.log("Calculated Total (Distributed + Overpaid): $" + calculatedTotal.toFixed(2));
  
  if (difference > 0.01) {
    Logger.log("⚠️ WARNING: Tip accounting mismatch!");
    Logger.log("  Difference: $" + difference.toFixed(2));
    Logger.log("  This indicates tips are being double-counted or lost!");
  } else {
    Logger.log("✓ Tips balance correctly");
  }
  
  return results;
}
    

/**
 * Find the closest booking to a transaction time
 * Uses time proximity - doesn't care about customer identity
 */
function findClosestBooking(transaction, bookings) {
  if (bookings.length === 0) return null;
  
  var transTime = transaction.dateTime.getTime();
  var closestBooking = null;
  var smallestTimeDiff = Infinity;
  
  // Look for bookings where transaction happened during or near the booking
  for (var i = 0; i < bookings.length; i++) {
    var booking = bookings[i];
    var bookingStart = booking.startTime.getTime();
    var bookingEnd = booking.endTime.getTime();
    
    // Check if transaction is during the booking
    if (transTime >= bookingStart && transTime <= bookingEnd) {
      // Transaction during booking = perfect match
      return booking;
    }
    
    // Calculate time difference (how far is transaction from this booking?)
    var timeDiff;
    if (transTime < bookingStart) {
      // Transaction before booking
      timeDiff = bookingStart - transTime;
    } else {
      // Transaction after booking
      timeDiff = transTime - bookingEnd;
    }
    
    // Keep track of closest booking (within 3 hours)
    var threeHoursMs = 3 * 60 * 60 * 1000;
    if (timeDiff < threeHoursMs && timeDiff < smallestTimeDiff) {
      smallestTimeDiff = timeDiff;
      closestBooking = booking;
    }
  }
  
  return closestBooking;
}

/**
 * Find bookings that match a transaction (by email and approximate time)
 */
function findMatchingBookingsWithLogging(transaction, bookings, shouldLog) {
  var matches = [];
  
  var transDate = new Date(transaction.dateTime);
  transDate.setHours(0, 0, 0, 0);
  
  if (shouldLog) {
    Logger.log("=== Matching transaction: " + transaction.transactionId + " ===");
    Logger.log("  Trans Email: '" + transaction.customerEmail + "'");
    Logger.log("  Trans Name: '" + transaction.customerName + "'");
    Logger.log("  Trans Date: " + formatDate(transDate));
  }
  
  // Strategy 1: Match by email on the same day (most common case)
  if (transaction.customerEmail) {
    if (shouldLog) Logger.log("  Trying email match...");
    var emailMatchCount = 0;
    
    for (var i = 0; i < bookings.length; i++) {
      var booking = bookings[i];
      
      if (booking.email === transaction.customerEmail) {
        var bookingDate = new Date(booking.startTime);
        bookingDate.setHours(0, 0, 0, 0);
        
        if (shouldLog) Logger.log("    Found email match: " + booking.email + " on " + formatDate(bookingDate));
        
        // Same email, same day = match
        if (bookingDate.getTime() === transDate.getTime()) {
          if (shouldLog) Logger.log("    ✓ Date matches! Adding booking.");
          matches.push(booking);
        } else {
          if (shouldLog) Logger.log("    ✗ Date doesn't match: " + formatDate(bookingDate) + " vs " + formatDate(transDate));
        }
        emailMatchCount++;
      }
    }
    
    if (shouldLog) Logger.log("  Found " + emailMatchCount + " bookings with matching email");
    
    if (matches.length > 0) {
      if (shouldLog) Logger.log("  ✓ Returning " + matches.length + " matched booking(s)");
      return matches;
    }
  } else {
    if (shouldLog) Logger.log("  No email on transaction, skipping email match");
  }
  
  // Strategy 2: Match by customer name on the same day
  if (transaction.customerName) {
    if (shouldLog) Logger.log("  Trying name match...");
    var transNameLower = transaction.customerName.toLowerCase().trim();
    
    for (var i = 0; i < bookings.length; i++) {
      var booking = bookings[i];
      var bookingDate = new Date(booking.startTime);
      bookingDate.setHours(0, 0, 0, 0);
      
      // Same day only
      if (bookingDate.getTime() === transDate.getTime()) {
        var bookingName = (booking.firstName + " " + booking.lastName).toLowerCase().trim();
        
        // Check if names match (full match or partial)
        if (bookingName && transNameLower) {
          // Split names into parts
          var transNameParts = transNameLower.split(/\s+/);
          var bookingNameParts = bookingName.split(/\s+/);
          
          // Check if any part matches
          var hasMatch = false;
          for (var j = 0; j < transNameParts.length; j++) {
            for (var k = 0; k < bookingNameParts.length; k++) {
              if (transNameParts[j].length >= 3 && bookingNameParts[k].length >= 3) {
                if (transNameParts[j] === bookingNameParts[k] || 
                    transNameParts[j].indexOf(bookingNameParts[k]) >= 0 ||
                    bookingNameParts[k].indexOf(transNameParts[j]) >= 0) {
                  hasMatch = true;
                  break;
                }
              }
            }
            if (hasMatch) break;
          }
          
          if (hasMatch) {
            if (shouldLog) Logger.log("    ✓ Name match: '" + transNameLower + "' matches '" + bookingName + "'");
            matches.push(booking);
          }
        }
      }
    }
    
    if (matches.length > 0) {
      if (shouldLog) Logger.log("  ✓ Returning " + matches.length + " matched booking(s) by name");
      return matches;
    }
  } else {
    if (shouldLog) Logger.log("  No name on transaction, skipping name match");
  }
  
  // No match found
  if (shouldLog) Logger.log("  ✗ NO MATCH FOUND - will split evenly among clocked-in employees");
  
  return matches;
}

/**
 * Find employees who worked during a specific time period and calculate overlap
 */
function findWorkingEmployees(staffShifts, bookingStart, bookingEnd) {
  var workingEmployees = [];
  
  var bookingStartTime = bookingStart.getTime();
  var bookingEndTime = bookingEnd.getTime();
  var bookingDuration = bookingEndTime - bookingStartTime;
  
  for (var i = 0; i < staffShifts.length; i++) {
    var shift = staffShifts[i];
    var shiftStartTime = shift.clockIn.getTime();
    var shiftEndTime = shift.clockOut.getTime();
    
    // Calculate overlap
    var overlapStart = Math.max(bookingStartTime, shiftStartTime);
    var overlapEnd = Math.min(bookingEndTime, shiftEndTime);
    
    if (overlapStart < overlapEnd) {
      // There is overlap
      var overlapDuration = overlapEnd - overlapStart;
      var overlapPercent = (overlapDuration / bookingDuration) * 100;
      
      workingEmployees.push({
        fullName: shift.fullName,
        firstName: shift.firstName,
        lastName: shift.lastName,
        jobTitle: shift.jobTitle,
        clockIn: shift.clockIn,
        clockOut: shift.clockOut,
        overlapPercent: overlapPercent,
        isTipEligible: shift.isTipEligible
      });
    }
  }
  
  return workingEmployees;
}

/**
 * Find employees who were clocked in at a specific point in time
 */
function findEmployeesAtTime(staffShifts, specificTime) {
  var clockedInEmployees = [];
  var timeMs = specificTime.getTime();
  
  for (var i = 0; i < staffShifts.length; i++) {
    var shift = staffShifts[i];
    var shiftStartTime = shift.clockIn.getTime();
    var shiftEndTime = shift.clockOut.getTime();
    
    // Check if the specific time falls within this shift
    if (timeMs >= shiftStartTime && timeMs <= shiftEndTime) {
      clockedInEmployees.push({
        fullName: shift.fullName,
        firstName: shift.firstName,
        lastName: shift.lastName,
        jobTitle: shift.jobTitle,
        clockIn: shift.clockIn,
        clockOut: shift.clockOut,
        isTipEligible: shift.isTipEligible
      });
    }
  }
  
  return clockedInEmployees;
}

// ============================================
// DISPLAY FUNCTIONS
// ============================================

/**
 * Display tip distribution results in formatted sheets
 */
function displayTipResults(results, payPeriod) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or clear the detailed breakdown sheet
  var detailSheet = ss.getSheetByName("Tip Distribution Details");
  if (detailSheet) {
    detailSheet.clear();
  } else {
    detailSheet = ss.insertSheet("Tip Distribution Details");
  }
  
  // Create or clear the summary sheet
  var summarySheet = ss.getSheetByName("Tip Distribution Summary");
  if (summarySheet) {
    summarySheet.clear();
  } else {
    summarySheet = ss.insertSheet("Tip Distribution Summary");
  }
  
  // === BUILD DETAILED BREAKDOWN ===
  displayDetailedBreakdown(detailSheet, results, payPeriod);
  
  // === BUILD SUMMARY ===
  displaySummary(summarySheet, results, payPeriod);
  
  // Activate summary sheet
  ss.setActiveSheet(summarySheet);
  
  // Show completion message
  var payPeriodStr = formatDate(payPeriod.start) + " to " + formatDate(payPeriod.end);
  var message = '✅ Tip Distribution Complete!\n\n';
  message += 'Pay Period: ' + payPeriodStr + '\n\n';
  message += 'Total Tips Processed: $' + results.totalTipsProcessed.toFixed(2) + '\n';
  message += 'Total Tips Distributed: $' + results.totalTipsDistributed.toFixed(2) + '\n';
  message += 'Overpaid (Not Distributed): $' + results.overpaid.toFixed(2) + '\n\n';
  message += 'Check "Tip Distribution Summary" for payroll entry\n';
  message += 'Check "Tip Distribution Details" for full breakdown';
  
  SpreadsheetApp.getUi().alert('Tip Distribution', message, SpreadsheetApp.getUi().ButtonSet.OK);
}

/**
 * Display detailed breakdown sheet
 */
function displayDetailedBreakdown(sheet, results, payPeriod) {
  try {
    Logger.log("Building detailed breakdown with " + results.details.length + " transactions");
    
    var currentRow = 1;
    
    // Header
    sheet.getRange("A1:H1").merge();
    sheet.getRange("A1").setValue("💰 TIP DISTRIBUTION - DETAILED BREAKDOWN");
    sheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A1").setBackground("#34A853").setFontColor("white");
    currentRow++;
    
    sheet.getRange("A2:H2").merge();
    var payPeriodStr = "Pay Period: " + formatDate(payPeriod.start) + " to " + formatDate(payPeriod.end) + " | Generated: " + new Date();
    sheet.getRange("A2").setValue(payPeriodStr);
    sheet.getRange("A2").setFontSize(10).setHorizontalAlignment("center").setBackground("#e6f4ea");
    currentRow += 2;
    
    // Build all data first, then write in one batch
    var allRows = [];
    
    // Process each transaction detail
    for (var i = 0; i < results.details.length; i++) {
      var detail = results.details[i];
      
      // Transaction header row
      var headerText = "Transaction: " + detail.transactionId + " | Customer: " + detail.customerName + " | Tip: $" + detail.tip.toFixed(2);
      if (detail.tip < 0) {
        headerText += " (REFUND/VOID)";
      }
      allRows.push([headerText, "", "", "", "", "", "", ""]);
      
      // Booking info row
      var bookingInfo = "Booking: " + formatDateTime(detail.bookingStart) + " to " + formatDateTime(detail.bookingEnd);
      allRows.push([bookingInfo, "", "", "", "", "", "", ""]);
      
      // Employee distribution
      if (detail.employeeDistribution.length > 0) {
        // Headers
        allRows.push(["Employee", "Clock In", "Clock Out", "Overlap %", "Tip Amount", "Notes", "", ""]);
        
        for (var j = 0; j < detail.employeeDistribution.length; j++) {
          var emp = detail.employeeDistribution[j];
          allRows.push([
            emp.employeeName,
            emp.clockIn ? formatDateTime(emp.clockIn) : "",
            emp.clockOut ? formatDateTime(emp.clockOut) : "",
            emp.overlapPercent.toFixed(1) + "%",
            "$" + emp.tipAmount.toFixed(2),
            emp.note || "",
            "",
            ""
          ]);
        }
      }
      
      // Overpaid amount if any
      if (detail.overpaid > 0) {
        allRows.push([
          "⚠️ OVERPAID (Not Distributed):",
          "",
          "$" + detail.overpaid.toFixed(2),
          detail.reason,
          "",
          "",
          "",
          ""
        ]);
      }
      
      // Blank row between transactions
      allRows.push(["", "", "", "", "", "", "", ""]);
    }
    
    // Write all data at once
    if (allRows.length > 0) {
      Logger.log("Writing " + allRows.length + " rows to detail sheet");
      sheet.getRange(currentRow, 1, allRows.length, 8).setValues(allRows);
    }
    
    // Auto-resize columns
    for (var col = 1; col <= 8; col++) {
      sheet.autoResizeColumn(col);
    }
    
    Logger.log("Detail breakdown complete");
    
  } catch (error) {
    Logger.log("ERROR in displayDetailedBreakdown: " + error.toString());
    throw error;
  }
}

/**
 * Display summary sheet for payroll entry
 */
function displaySummary(sheet, results, payPeriod) {
  try {
    Logger.log("Building summary sheet");
    
    // Header
    sheet.getRange("A1:E1").merge();
    sheet.getRange("A1").setValue("💰 TIP DISTRIBUTION SUMMARY - FOR PAYROLL");
    sheet.getRange("A1").setFontSize(16).setFontWeight("bold").setHorizontalAlignment("center");
    sheet.getRange("A1").setBackground("#34A853").setFontColor("white");
    
    sheet.getRange("A2:E2").merge();
    var payPeriodStr = "Pay Period: " + formatDate(payPeriod.start) + " to " + formatDate(payPeriod.end) + " | Generated: " + new Date();
    sheet.getRange("A2").setValue(payPeriodStr);
    sheet.getRange("A2").setFontSize(10).setHorizontalAlignment("center").setBackground("#e6f4ea");
    
    var currentRow = 4;
    
    // Calculate wages from shifts
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var staffSheet = ss.getSheetByName("Staff Timecards");
    var staffData = staffSheet.getDataRange().getValues();
    var staffHeaders = staffData[0];
    
    var firstNameCol = staffHeaders.indexOf("First name");
    var lastNameCol = staffHeaders.indexOf("Last name");
    var totalPaidHoursCol = staffHeaders.indexOf("Total paid hours");
    
    var employeeWages = {};
    var employeeHours = {};
    
    for (var i = 1; i < staffData.length; i++) {
      var firstName = staffData[i][firstNameCol] || "";
      var lastName = staffData[i][lastNameCol] || "";
      var fullName = (firstName + " " + lastName).trim();
      
      if (!fullName) continue;
      
      var hours = parseFloat(staffData[i][totalPaidHoursCol]) || 0;
      
      if (!employeeHours[fullName]) {
        employeeHours[fullName] = 0;
      }
      employeeHours[fullName] += hours;
    }
    
    // Calculate wages
    var HOURLY_RATE = 17.50;
    var BRONSON_SALARY = 2308.00;
    
    for (var empName in employeeHours) {
      if (empName.toLowerCase() === "bronson roberts") {
        employeeWages[empName] = BRONSON_SALARY; // Salary
      } else {
        employeeWages[empName] = employeeHours[empName] * HOURLY_RATE;
      }
    }
    
    // PAYROLL SUMMARY TABLE
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue("📋 COMPLETE PAYROLL SUMMARY");
    sheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#4285F4").setFontColor("white");
    currentRow++;
    
    var headers = ["Employee Name", "Hours", "Wages", "Tips", "Total Pay"];
    sheet.getRange(currentRow, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(currentRow, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8E8E8");
    currentRow++;
    
    // Sort employees by name
    var employeeList = [];
    var allEmployees = {};
    
    // Combine all employees (from hours and tips)
    for (var empName in employeeHours) {
      allEmployees[empName] = true;
    }
    for (var empName in results.summary) {
      allEmployees[empName] = true;
    }
    
    for (var empName in allEmployees) {
      var hours = employeeHours[empName] || 0;
      var wages = employeeWages[empName] || 0;
      var tips = results.summary[empName] || 0;
      var totalPay = wages + tips;
      
      employeeList.push({
        name: empName,
        hours: hours,
        wages: wages,
        tips: tips,
        totalPay: totalPay
      });
    }
    
    employeeList.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
    
    // Build employee data array
    var employeeData = [];
    var totalHours = 0;
    var totalWages = 0;
    var totalTips = 0;
    var totalPayroll = 0;
    
    for (var i = 0; i < employeeList.length; i++) {
      var emp = employeeList[i];
      var hoursDisplay = emp.name.toLowerCase() === "bronson roberts" ? "Salary" : emp.hours.toFixed(2);
      
      employeeData.push([
        emp.name,
        hoursDisplay,
        emp.wages,
        emp.tips,
        emp.totalPay
      ]);
      
      if (emp.name.toLowerCase() !== "bronson roberts") {
        totalHours += emp.hours;
      }
      totalWages += emp.wages;
      totalTips += emp.tips;
      totalPayroll += emp.totalPay;
    }
    
    // Write employee payroll
    if (employeeData.length > 0) {
      Logger.log("Writing " + employeeData.length + " employee payroll records");
      sheet.getRange(currentRow, 1, employeeData.length, 5).setValues(employeeData);
      sheet.getRange(currentRow, 3, employeeData.length, 3).setNumberFormat("$#,##0.00");
      currentRow += employeeData.length;
      
      // Totals row
      sheet.getRange(currentRow, 1).setValue("TOTALS:");
      sheet.getRange(currentRow, 1).setFontWeight("bold");
      sheet.getRange(currentRow, 2).setValue(totalHours.toFixed(2) + " hrs");
      sheet.getRange(currentRow, 3).setValue(totalWages);
      sheet.getRange(currentRow, 4).setValue(totalTips);
      sheet.getRange(currentRow, 5).setValue(totalPayroll);
      sheet.getRange(currentRow, 3, 1, 3).setNumberFormat("$#,##0.00");
      sheet.getRange(currentRow, 1, 1, 5).setFontWeight("bold").setBackground("#d9ead3");
      currentRow++;
      
      // Add borders
      sheet.getRange(currentRow - employeeData.length - 2, 1, employeeData.length + 2, 5).setBorder(true, true, true, true, true, true);
    }
    
    currentRow += 2;
    
    // TIP BREAKDOWN (for reference)
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue("💵 TIP BREAKDOWN DETAIL");
    sheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#FF6D00").setFontColor("white");
    currentRow++;
    
    var tipHeaders = ["Employee Name", "Total Tips", "", "", ""];
    sheet.getRange(currentRow, 1, 1, tipHeaders.length).setValues([tipHeaders]);
    sheet.getRange(currentRow, 1, 1, tipHeaders.length).setFontWeight("bold").setBackground("#E8E8E8");
    currentRow++;
    
    // Tip-only list (sorted)
    var tipList = [];
    for (var empName in results.summary) {
      tipList.push({
        name: empName,
        amount: results.summary[empName]
      });
    }
    tipList.sort(function(a, b) {
      return a.name.localeCompare(b.name);
    });
    
    var tipData = [];
    for (var i = 0; i < tipList.length; i++) {
      tipData.push([
        tipList[i].name,
        tipList[i].amount,
        "",
        "",
        ""
      ]);
    }
    
    if (tipData.length > 0) {
      sheet.getRange(currentRow, 1, tipData.length, 5).setValues(tipData);
      sheet.getRange(currentRow, 2, tipData.length, 1).setNumberFormat("$#,##0.00");
      sheet.getRange(currentRow - 1, 1, tipData.length + 1, 2).setBorder(true, true, true, true, true, true);
      currentRow += tipData.length;
    }
    
    currentRow += 2;
    
    // Totals section
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue("📊 TOTALS & RECONCILIATION");
    sheet.getRange(currentRow, 1).setFontWeight("bold").setFontSize(12).setBackground("#9C27B0").setFontColor("white");
    currentRow++;
    
    var totalsData = [
      ["Total Wages (incl. Bronson salary)", "$" + totalWages.toFixed(2), "", "", ""],
      ["Total Tips Distributed", "$" + results.totalTipsDistributed.toFixed(2), "", "", ""],
      ["Total Payroll Cost", "$" + totalPayroll.toFixed(2), "", "", ""],
      ["", "", "", "", ""],
      ["Total Tips Processed", "$" + results.totalTipsProcessed.toFixed(2), "", "", ""],
      ["Overpaid (Not Distributed)", "$" + results.overpaid.toFixed(2), "", "", ""],
      ["✅ Tips Should Balance:", "$" + (results.totalTipsDistributed + results.overpaid).toFixed(2), "", "", ""]
    ];
    
    sheet.getRange(currentRow, 1, totalsData.length, 5).setValues(totalsData);
    sheet.getRange(currentRow, 1, totalsData.length, 1).setBackground("#f3e5f5");
    sheet.getRange(currentRow + totalsData.length - 1, 1, 1, 2).setFontWeight("bold").setBackground("#d9ead3");
    currentRow += totalsData.length + 2;
    
    // Notes section
    sheet.getRange(currentRow, 1, 1, 5).merge();
    sheet.getRange(currentRow, 1).setValue("📝 NOTES");
    sheet.getRange(currentRow, 1).setFontWeight("bold").setBackground("#e8f0fe");
    currentRow++;
    
    var notes = [
      ["• Hourly employees earn $17.50/hour", "", "", "", ""],
      ["• Bronson Roberts is salaried at $2,308.00 per pay period", "", "", "", ""],
      ["• Tips are distributed proportionally based on employee work overlap with customer booking times", "", "", "", ""],
      ["• When no booking is found, tips are split evenly among employees clocked in at payment time", "", "", "", ""],
      ["• Bronson Roberts is NOT tip-eligible - tips are distributed to other working employees", "", "", "", ""],
      ["• If only Bronson worked, tips go to 'Overpaid' and are not distributed", "", "", "", ""],
      ["• Check 'Tip Distribution Details' sheet for full breakdown of each transaction", "", "", "", ""]
    ];
    
    sheet.getRange(currentRow, 1, notes.length, 5).setValues(notes);
    for (var i = 0; i < notes.length; i++) {
      sheet.getRange(currentRow + i, 1, 1, 5).merge();
      sheet.getRange(currentRow + i, 1).setWrap(true);
    }
    
    // Auto-resize columns
    sheet.setColumnWidth(1, 300);
    sheet.setColumnWidth(2, 100);
    sheet.setColumnWidth(3, 120);
    sheet.setColumnWidth(4, 120);
    sheet.setColumnWidth(5, 120);
    
    Logger.log("Summary complete");
    
  } catch (error) {
    Logger.log("ERROR in displaySummary: " + error.toString());
    throw error;
  }
}

// ============================================
// HELPER FUNCTIONS
// ============================================

/**
 * Parse date/time from Square format
 */
function parseDateTimeFromSquare(date, time) {
  if (!date) return null;
  
  var dateObj = new Date(date);
  
  if (!time) {
    return dateObj;
  }
  
  // Parse time string (e.g., "4:54:35 PM PDT")
  var timeStr = String(time);
  var match = timeStr.match(/(\d+):(\d+):(\d+)\s*(AM|PM)/i);
  
  if (match) {
    var hours = parseInt(match[1]);
    var minutes = parseInt(match[2]);
    var seconds = parseInt(match[3]);
    var meridiem = match[4].toUpperCase();
    
    if (meridiem === "PM" && hours !== 12) {
      hours += 12;
    } else if (meridiem === "AM" && hours === 12) {
      hours = 0;
    }
    
    dateObj.setHours(hours, minutes, seconds, 0);
  }
  
  return dateObj;
}

/**
 * Normalize email address for comparison
 */
function normalizeEmail(email) {
  if (!email) return "";
  return String(email).toLowerCase().trim();
}

/**
 * Combine date and time into single Date object
 */
function combineDateAndTime(date, time) {
  if (!date) return null;
  
  var dateObj = new Date(date);
  
  if (!time) {
    return dateObj;
  }
  
  // Handle different time formats
  if (typeof time === 'string') {
    var match = time.match(/(\d+):(\d+):(\d+)/);
    if (match) {
      dateObj.setHours(parseInt(match[1]), parseInt(match[2]), parseInt(match[3]), 0);
    }
  } else if (time instanceof Date) {
    dateObj.setHours(time.getHours(), time.getMinutes(), time.getSeconds(), 0);
  }
  
  return dateObj;
}

/**
 * Format date/time for display
 */
function formatDateTime(dateTime) {
  if (!dateTime) return "";
  return Utilities.formatDate(dateTime, Session.getScriptTimeZone(), "MM/dd/yyyy hh:mm a");
}

/**
 * Format date only for display
 */
function formatDate(date) {
  if (!date) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "MM/dd/yyyy");
}