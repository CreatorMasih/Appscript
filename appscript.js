function setup() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Ensure the Summary sheet exists
    if (!ss.getSheetByName('Summary')) {
      ss.insertSheet('Summary');
    }
    
    var summarySheet = ss.getSheetByName('Summary');
    
    // Set headers for Summary sheet
    summarySheet.clear();
    summarySheet.getRange('A1').setValue('Student Name');
    summarySheet.getRange('B1').setValue('Contact Number');
    summarySheet.getRange('C1').setValue('Number Validity');
    summarySheet.getRange('D1').setValue('Potential');
    summarySheet.getRange('E1').setValue('1st Call Status');
    summarySheet.getRange('F1').setValue('WA Message');
    summarySheet.getRange('G1').setValue('WhatsApp Link');
    summarySheet.getRange('H1').setValue('Additional Info');
    summarySheet.getRange('J1').setValue('Potential Summary');
    summarySheet.getRange('J2').setValue('Low');
    summarySheet.getRange('J3').setValue('Medium');
    summarySheet.getRange('J4').setValue('High');
    summarySheet.getRange('J5').setValue('Total');
  
    // Apply formatting to headers
    var headerRange = summarySheet.getRange('A1:J1');
    headerRange.setFontWeight('bold');
    headerRange.setBackground('#4285F4'); // Header color (blue)
    headerRange.setFontColor('#ffffff'); // Font color (white)
  }
  
  function validateAndSummarize() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var callStatusSheet = ss.getSheetByName('Call Status');
  
    if (!callStatusSheet) {
      Logger.log("Sheet 'Call Status' not found!");
      return;
    }
    
    var data = callStatusSheet.getDataRange().getValues();
    
    var summarySheet = ss.getSheetByName('Summary');
    summarySheet.clearContents(); // Clear previous summary data but keep headers
    
    var headers = data[0];
    var taskDateIndex = headers.indexOf('Task Date');
    
    // Add Task Date to headers if it exists
    if (taskDateIndex > -1) {
      summarySheet.getRange('I1').setValue('Task Date');
    }
    
    var rowCount = 2; // Start writing to summary sheet from row 2
    
    var lowCount = 0;
    var mediumCount = 0;
    var highCount = 0;
    
    for (var i = 1; i < data.length; i++) {
      var studentName = data[i][0];
      var contactNumber = data[i][1];
      var potential = data[i][3];
      var waMessage = data[i][4];
      var firstCallStatus = data[i][5];
      var taskDate = taskDateIndex > -1 ? data[i][taskDateIndex] : '';
      
      var numberValidity = validateContactNumber(contactNumber);
      var whatsappLink = createWhatsappLink(contactNumber); // Generate WhatsApp link
      
      // Populate summary sheet
      summarySheet.getRange('A' + rowCount).setValue(studentName);
      summarySheet.getRange('B' + rowCount).setValue(contactNumber);
      summarySheet.getRange('C' + rowCount).setValue(numberValidity);
      summarySheet.getRange('D' + rowCount).setValue(potential);
      summarySheet.getRange('E' + rowCount).setValue(firstCallStatus);
      summarySheet.getRange('F' + rowCount).setValue(waMessage);
      summarySheet.getRange('G' + rowCount).setValue(whatsappLink); // Set WhatsApp link
      
      if (taskDateIndex > -1) {
        summarySheet.getRange('I' + rowCount).setValue(taskDate);
      }
      
      // Conditional checks for additional information
      var additionalInfo = '';
      if (numberValidity === "Invalid Number") {
        additionalInfo += 'Invalid Contact Number; ';
      }
      
      if (potential.toLowerCase() === 'low') {
        additionalInfo += 'Low Potential - ' + firstCallStatus + '; ';
        lowCount++;
      } else if (potential.toLowerCase() === 'medium') {
        mediumCount++;
      } else if (potential.toLowerCase() === 'high') {
        highCount++;
      }
      
      if (waMessage.toLowerCase() !== 'sent') {
        additionalInfo += 'WA Message Not Sent - ' + firstCallStatus + '; ';
      }
      
      summarySheet.getRange(taskDateIndex > -1 ? 'H' : 'H' + rowCount).setValue(additionalInfo);
      
      // Apply color formatting
      summarySheet.getRange('A' + rowCount).setBackground('#D9EAD3'); // Light blue for Student Name
      summarySheet.getRange('B' + rowCount).setBackground(numberValidity === "Valid Number" ? '#006400' : '#8B0000'); // Green if valid, red if invalid for Contact Number
      summarySheet.getRange('D' + rowCount).setBackground(potential.toLowerCase() === 'low' ? '#F4CCCC' : (potential.toLowerCase() === 'medium' ? '#FFF2CC' : '#D9EAD3')); // Red for low, yellow for medium, green for high
      summarySheet.getRange('E' + rowCount).setBackground(potential.toLowerCase() === 'low' ? '#F4CCCC' : (potential.toLowerCase() === 'medium' ? '#FFF2CC' : '#D9EAD3')); // Same color as Potential for 1st Call Status
      summarySheet.getRange('F' + rowCount).setBackground(waMessage.toLowerCase() === 'sent' ? '#006400' : '#8B0000'); // Green if sent, red if not sent for WA Message
      summarySheet.getRange('H' + rowCount).setBackground(additionalInfo.includes('WA Message Not Sent') && potential.toLowerCase() === 'low' ? '#F4CCCC' : '#FFF2CC'); // Red if WA message not sent and potential is low, yellow otherwise
      
      rowCount++;
    }
    
    // Set potential summary counts
    summarySheet.getRange('J2').setValue("Low: " + lowCount);
    summarySheet.getRange('J3').setValue("Medium: " + mediumCount);
    summarySheet.getRange('J4').setValue("High: " + highCount);
    summarySheet.getRange('J5').setValue("Total: " + (lowCount + mediumCount + highCount)); // Total count
  }
  
  function validateContactNumber(contactNumber) {
    if (typeof contactNumber === 'undefined' || contactNumber === null) {
      return "Invalid Number";
    }
    var contactNumberStr = contactNumber.toString();
    var digits = contactNumberStr.replace(/\D/g, '');
    return (digits.length === 10) ? "Valid Number" : "Invalid Number";
  }
  
  function createWhatsappLink(contactNumber) {
    if (typeof contactNumber === 'undefined' || contactNumber === null) {
      return '';
    }
    var contactNumberStr = contactNumber.toString();
    var digits = contactNumberStr.replace(/\D/g, '');
    return 'https://wa.me/' + digits;
  }
  
  function setUpTrigger() {
    ScriptApp.newTrigger('validateAndSummarize')
      .timeBased()
      .everyDays(1)
      .atHour(1)
      .create();
  }
  
  function initialize() {
    setup();
    validateAndSummarize();
  }
  