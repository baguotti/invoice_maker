/**
 * @OnlyCurrentDoc
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Invoice')
      .addItem('Generate from Selection', 'generateFromSelection')
      .addItem('Generate Quote from Selection', 'generateQuoteFromSelection')
      .addItem('Manage Clients...', 'showManageClientsDialog')
      .addToUi();
}

// Called from Setup.html to get saved clients
function getSavedClients() {
  var props = PropertiesService.getDocumentProperties();
  var clientsStr = props.getProperty('saved_clients');
  if (clientsStr) {
    return JSON.parse(clientsStr);
  }
  return {};
}

// Called from Setup.html to save a new client
function saveClient(nickname, details) {
  var props = PropertiesService.getDocumentProperties();
  var clientsStr = props.getProperty('saved_clients');
  var clients = clientsStr ? JSON.parse(clientsStr) : {};
  clients[nickname] = details;
  props.setProperty('saved_clients', JSON.stringify(clients));
  return clients;
}

function showManageClientsDialog() {
  var html = HtmlService.createHtmlOutputFromFile('ManageClients')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Clients');
}

function deleteClient(nickname) {
  var props = PropertiesService.getDocumentProperties();
  var clientsStr = props.getProperty('saved_clients');
  if (clientsStr) {
    var clients = JSON.parse(clientsStr);
    delete clients[nickname];
    props.setProperty('saved_clients', JSON.stringify(clients));
    return clients;
  }
  return {};
}

function renameClient(oldNickname, newNickname, details) {
  var props = PropertiesService.getDocumentProperties();
  var clientsStr = props.getProperty('saved_clients');
  var clients = clientsStr ? JSON.parse(clientsStr) : {};
  
  if (oldNickname && oldNickname !== newNickname) {
    delete clients[oldNickname];
  }
  clients[newNickname] = details;
  props.setProperty('saved_clients', JSON.stringify(clients));
  return clients;
}

function getWeekRangeLabel(date) {
  var d = new Date(date);
  d.setHours(0,0,0,0);
  
  // Get Monday
  var day = d.getDay();
  var diff = d.getDate() - day + (day === 0 ? -6 : 1);
  var monday = new Date(d.setDate(diff));
  
  // Get Sunday
  var sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  var options = { month: 'short', day: 'numeric' };
  return monday.toLocaleDateString('en-US', options) + ' - ' + sunday.toLocaleDateString('en-US', options);
}

// Helper to format currency
function formatCurrency(num) {
  return Number(num).toLocaleString('en-GB', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function generateFromSelection() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangeList = sheet.getActiveRangeList();
  var ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];
  
  var data = [];
  for (var k = 0; k < ranges.length; k++) {
    var r = ranges[k];
    data = data.concat(sheet.getRange(r.getRow(), 1, r.getNumRows(), 4).getValues());
  }
  
  var projectStr = "";
  // Find the first valid project string in the selection
  // Format: #0000 DESCRIPTION @CLIENT NICKNAME
  for (var i = 0; i < data.length; i++) {
     for (var j = 0; j < data[i].length; j++) {
       var cellVal = String(data[i][j]).trim();
       if (cellVal.match(/^#\S+\s+.*?\s+@/)) {
          projectStr = cellVal;
          break;
       }
     }
     if (projectStr) break;
  }
  
  if (!projectStr) {
     SpreadsheetApp.getUi().alert("Could not find a project description in the highlighted cells.");
     return;
  }
  
  // Format: #0000 DESCRIPTION @CLIENT NICKNAME
  var match = projectStr.match(/#(\S+)\s+(.*?)\s+@(.*)/);
  if (!match) {
     SpreadsheetApp.getUi().alert("Invalid project format.\nExpected: #0000 DESCRIPTION @CLIENT NICKNAME\nFound: " + projectStr);
     return;
  }
  
  var invoiceNum = match[1];
  var forText = match[2].trim();
  var clientNickname = match[3].trim();
  
  var props = PropertiesService.getDocumentProperties();
  var clientsStr = props.getProperty('saved_clients');
  var clients = clientsStr ? JSON.parse(clientsStr) : {};
  
  var clientDetails = clients[clientNickname];
  if (!clientDetails) {
     SpreadsheetApp.getUi().alert("Client '" + clientNickname + "' not found. Please add them via 'Manage Clients...' first.");
     return;
  }
  
  generateInvoiceHTML(clientDetails, invoiceNum, forText, clientNickname);
}

// Called from generateFromSelection to generate the actual invoice
function generateInvoiceHTML(clientDetails, invoiceNum, forText, clientNickname) {
  var sheet = SpreadsheetApp.getActiveSheet();
  
  var rangeList = sheet.getActiveRangeList();
  var ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];
  var data = [];
  var lowestRow = 0;
  
  for (var k = 0; k < ranges.length; k++) {
    var r = ranges[k];
    var rStart = r.getRow();
    var rNum = r.getNumRows();
    data = data.concat(sheet.getRange(rStart, 1, rNum, 4).getValues());
    
    var rEnd = rStart + rNum - 1;
    if (rEnd > lowestRow) lowestRow = rEnd;
  }
  
  var weeksMap = {};
  var totalDays = 0;
  var totalFee = 0;
  var currency = "£"; // Currency symbol based on prompt example
  
  var minDate = null;
  var maxDate = null;
  
  var startIndex = 0;
  if (data.length > 0 && String(data[0][0]).trim().toUpperCase() === "DATE") {
    startIndex = 1;
  }
  
  for (var i = startIndex; i < data.length; i++) {
    var row = data[i];
    var dateVal = row[0];
    var project = row[1];
    var type = row[2];
    var rateVal = row[3];
    
    // Check if type is "D" (case insensitive)
    if (String(type).trim().toUpperCase() !== "D") continue;
    
    // Parse date
    var d = new Date(dateVal);
    if (isNaN(d.getTime())) continue; // Skip invalid dates
    
    if (!minDate || d < minDate) minDate = new Date(d);
    if (!maxDate || d > maxDate) maxDate = new Date(d);
    
    var weekLabel = getWeekRangeLabel(d);
    
    // Parse rate
    var rate = 0;
    if (typeof rateVal === 'number') {
      rate = rateVal;
    } else if (typeof rateVal === 'string') {
      // extract numbers and decimals
      var matches = rateVal.match(/[0-9.]+/g);
      if (matches) {
        rate = parseFloat(matches.join(''));
      }
    }
    
    var groupKey = weekLabel + '_' + rate;
    
    if (!weeksMap[groupKey]) {
      weeksMap[groupKey] = {
        label: weekLabel,
        originalWeekLabel: weekLabel,
        days: 0,
        dates: [],
        rate: rate,
        total: 0,
        projects: new Set(),
        sortKey: d.getTime()
      };
    }
    
    weeksMap[groupKey].days += 1;
    
    // Format the date like "10/08"
    var dateStr = ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2);
    if (!weeksMap[groupKey].dates.includes(dateStr)) {
        weeksMap[groupKey].dates.push(dateStr);
    }

    var displayProject = project;
    if (project) {
      var m = String(project).match(/#(\S+)\s+(.*?)\s+@(.*)/);
      if (m) displayProject = m[2].trim();
    }
    
    if (displayProject) weeksMap[groupKey].projects.add(String(displayProject).trim());
    weeksMap[groupKey].total += rate;
    
    totalDays += 1;
    totalFee += rate;
  }
  
  // Format data for template
  var groupedData = [];
  for (var key in weeksMap) {
    weeksMap[key].dates.sort();

    groupedData.push({
      originalWeekLabel: weeksMap[key].originalWeekLabel,
      datesWorked: weeksMap[key].dates.join(', '),
      days: weeksMap[key].days,
      rateDisplay: currency + formatCurrency(weeksMap[key].rate),
      total: weeksMap[key].total,
      totalDisplay: formatCurrency(weeksMap[key].total),
      projects: Array.from(weeksMap[key].projects).filter(Boolean).join(', '),
      sortKey: weeksMap[key].sortKey
    });
  }
  
  // Sort by week chronological order
  groupedData.sort(function(a, b) {
    return a.sortKey - b.sortKey;
  });
  
  // Assign "Week X" labels
  if (groupedData.length > 0) {
    var currentWeekNum = 1;
    var currentWeekLabel = groupedData[0].originalWeekLabel;
    for (var j = 0; j < groupedData.length; j++) {
      if (groupedData[j].originalWeekLabel !== currentWeekLabel) {
        currentWeekNum++;
        currentWeekLabel = groupedData[j].originalWeekLabel;
      }
      groupedData[j].weekLabel = "Week " + currentWeekNum;
    }
  }
  
  var intervalStr = "";
  if (minDate && maxDate) {
    var options = { month: 'short', day: 'numeric', year: 'numeric' };
    if (minDate.getTime() === maxDate.getTime()) {
      intervalStr = minDate.toLocaleDateString('en-GB', options);
    } else {
      intervalStr = minDate.toLocaleDateString('en-GB', options) + ' - ' + maxDate.toLocaleDateString('en-GB', options);
    }
  } else {
    intervalStr = "No dates selected";
  }
  var documentTitle = invoiceNum + " - Invoice Ric_" + forText;
  
  var template = HtmlService.createTemplateFromFile('Invoice');
  template.documentTitle = documentTitle;
  template.intervalStr = intervalStr;
  template.groupedData = groupedData;
  template.totalDays = totalDays;
  template.totalFee = totalFee;
  template.totalFeeDisplay = formatCurrency(totalFee);
  template.currency = currency;
  template.clientDetails = clientDetails;
  template.invoiceNum = invoiceNum;
  template.forText = forText;
  template.lowestRow = lowestRow;
  template.clientNickname = clientNickname;
  
  var isUs = false;
  if (clientDetails) {
    var addrUpper = String(clientDetails).toUpperCase();
    var hasUs = /\b(US|USA|U\.S\.A\.|UNITED STATES|UNITED STATES OF AMERICA)\b/i.test(addrUpper);
    var hasUk = /\b(UK|U\.K\.|UNITED KINGDOM|GREAT BRITAIN|GB)\b/i.test(addrUpper);
    isUs = hasUs && !hasUk;
  }
  template.isUs = isUs;
  
  var now = new Date();
  var dateOptions = { day: 'numeric', month: 'long', year: 'numeric' };
  template.currentDate = now.toLocaleDateString('en-GB', dateOptions);
  
  var htmlOutput = template.evaluate()
    .setWidth(850)
    .setHeight(650)
    .setTitle(documentTitle);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Invoice');
}

function generateQuoteFromSelection() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rangeList = sheet.getActiveRangeList();
  var ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];
  
  var data = [];
  var displayData = [];
  
  for (var k = 0; k < ranges.length; k++) {
    var r = ranges[k];
    var startRow = r.getRow();
    var numRows = r.getNumRows();
    data = data.concat(sheet.getRange(startRow, 1, numRows, 6).getValues());
    displayData = displayData.concat(sheet.getRange(startRow, 1, numRows, 6).getDisplayValues());
  }
  
  // Detect currency from display values
  var currencySymbol = "£"; // Default
  for (var r = 0; r < displayData.length; r++) {
    var rowText = displayData[r].join(" ");
    if (rowText.indexOf('$') !== -1) {
      currencySymbol = "$";
      break;
    }
    if (rowText.indexOf('£') !== -1) {
      currencySymbol = "£";
      break;
    }
  }

  
  var quoteDate = "";
  var projectDesc = "";
  var dayRate = "";
  var items = [];
  var totalAmount = "";
  var totalDays = 0;
  
  var headerFound = false;
  
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    var colA = String(row[0]).trim();
    var colB = String(row[1]).trim();
    var colC = String(row[2]).trim();
    var colD = String(row[3]).trim();
    var colE = String(row[4]).trim();
    var colF = String(row[5]).trim();
    
    // Parse top header section
    if (colA.toUpperCase() === "DATE") {
      var dateCell = row[1];
      if (Object.prototype.toString.call(dateCell) === '[object Date]') {
         quoteDate = dateCell.toDateString();
      } else {
         quoteDate = colB;
         var match = quoteDate.match(/(?:Quote\s*)?(\d+)\/(\d+)\/(\d+)/i);
         if (match) {
            var d = parseInt(match[1], 10);
            var m = parseInt(match[2], 10);
            var y = parseInt(match[3], 10);
            if (y < 100) y += 2000;
            var dateObj = new Date(y, m - 1, d);
            quoteDate = dateObj.toDateString(); 
         }
      }
    }
    if (colA.toUpperCase() === "PROJECT DESCRIPTION") {
      projectDesc = colB;
    }
    
    // Search for "DAY RATE" in any column
    var dayRateColIndex = -1;
    for (var col = 0; col < row.length; col++) {
       if (String(row[col]).trim().toUpperCase() === "DAY RATE") {
          dayRateColIndex = col;
          break;
       }
    }
    if (dayRateColIndex !== -1) {
       if (i + 1 < data.length) {
          var rawDayRate = String(data[i+1][dayRateColIndex]).trim();
          if (rawDayRate !== "") dayRate = rawDayRate;
       }
    }
    
    if (colB.toUpperCase() === "ITEM" && colC.toUpperCase() === "DESCRIPTION" && colD.toUpperCase() === "DAYS") {
      headerFound = true;
      continue;
    }
    
    if (headerFound) {
      if (colA.toUpperCase() === "TOTAL" || colB.toUpperCase() === "TOTAL" || colC.toUpperCase() === "TOTAL" || colD.toUpperCase() === "TOTAL") {
         totalAmount = colF;
         if (!totalAmount && colD) totalAmount = colD; // fallback
         break;
      }
      
      // Look for a valid item
      if (colB !== "" || colC !== "") {
         items.push({
            item: colB,
            description: colC,
            days: colD,
            totalDisplay: colF
         });
         var dVal = parseFloat(colD);
         if (!isNaN(dVal)) totalDays += dVal;
      }
    }
  }
  
  if (!quoteDate && !projectDesc && items.length === 0) {
     SpreadsheetApp.getUi().alert("Could not find quote data. Please highlight the entire quote block (from DATE to Total).");
     return;
  }
  
  function customFormatCurrency(val) {
    var str = String(val).trim();
    if (!str) return "";
    
    var matches = str.match(/[0-9.,]+/g);
    if (matches) {
       var numStr = matches.join('').replace(/,/g, '');
       var num = parseFloat(numStr);
       if (!isNaN(num)) {
          return num.toLocaleString('en-GB', {minimumFractionDigits: 0, maximumFractionDigits: 2});
       }
    }
    return str;
  }


  
  var template = HtmlService.createTemplateFromFile('Quote');
  template.documentTitle = "Quote - " + projectDesc;
  template.quoteDate = quoteDate;
  template.clientDetails = projectDesc; // Using project description for "For:" field
  template.dayRate = customFormatCurrency(dayRate);


  
  // Format all item amounts
  for (var j = 0; j < items.length; j++) {
     items[j].totalDisplay = customFormatCurrency(items[j].totalDisplay);


  }
  
  template.groupedData = items;
  template.totalDays = totalDays;
  template.totalFeeDisplay = customFormatCurrency(totalAmount);
  template.currency = currencySymbol;

  
  var htmlOutput = template.evaluate()
    .setWidth(850)
    .setHeight(650)
    .setTitle(template.documentTitle);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Quote');
}

/**
 * Saves a base64-encoded PDF file to the specific Google Drive folder.
 * Called from client-side JS in Invoice.html.
 *
 * @param {string} base64Data
 * @param {string} fileName
 * @param {string} invoiceNum
 * @param {number} lowestRow
 * @param {boolean} sendEmail
 * @param {string} base64Data
 * @param {string} fileName
 * @param {string} invoiceNum
 * @param {number} lowestRow
 * @param {boolean} sendEmail
 * @param {string} forText
 * @return {string} The URL of the saved file.
 */
function savePdfToDrive(base64Data, fileName, invoiceNum, lowestRow, sendEmail, forText) {
  try {
    var folderId = "1_SZuoOUNIw9UIFp-I6DczLQ1l7oJ9B67";
    var folder = DriveApp.getFolderById(folderId);
    var contentType = "application/pdf";
    var decoded = Utilities.base64Decode(base64Data);
    var blob = Utilities.newBlob(decoded, contentType, fileName);
    var file = folder.createFile(blob);
    
    // Auto-tag the sheet with the invoice number as a hyperlink to the PDF
    if (invoiceNum && lowestRow) {
      var sheet = SpreadsheetApp.getActiveSheet();
      var fileUrl = file.getUrl();
      
      var cleanInvoiceNum = String(invoiceNum).replace(/["']/g, "");
      var richText = SpreadsheetApp.newRichTextValue()
        .setText("#" + cleanInvoiceNum)
        .setLinkUrl(fileUrl)
        .build();
        
      sheet.getRange(lowestRow, 7).setRichTextValue(richText);
      sheet.getRange(lowestRow, 8).setValue(new Date()); // Column H is the generation date
      sheet.getRange(lowestRow, 9).setValue("N"); // Column I is "N"
    }
    
    // Send email with PDF attachment if requested
    if (sendEmail) {
      var cleanInvoiceNum = String(invoiceNum).replace(/["']/g, "");
      var cleanForText = String(forText || "").replace(/["']/g, "").trim();
      MailApp.sendEmail({
        to: "accounts@wildislandfilms.com, amy@wildislandfilms.com, penina@wildislandfilms.com",
        subject: "Ric invoice #" + cleanInvoiceNum,
        body: "Heya,\n\nPlease find attached invoice #" + cleanInvoiceNum + " for " + cleanForText + "\n\nBest,\nRic",
        attachments: [blob]
      });
    }
    
    return file.getUrl();
  } catch (e) {
    throw new Error("Failed to save PDF to Google Drive: " + e.toString());
  }
}

