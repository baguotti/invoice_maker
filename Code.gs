/**
 * @OnlyCurrentDoc
 */

const CONFIG = {
  DRIVE_FOLDER_ID: "1_SZuoOUNIw9UIFp-I6DczLQ1l7oJ9B67",
  DEFAULT_EMAIL_RECIPIENTS: "accounts@wildislandfilms.com, amy@wildislandfilms.com, penina@wildislandfilms.com"
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Generate Invoice')
      .addItem('Generate from Selection', 'generateFromSelection')
      .addItem('Generate Quote from Selection', 'generateQuoteFromSelection')
      .addItem('Manage Clients...', 'showManageClientsDialog')
      .addToUi();
}

/** Helper to get clients from properties */
function _getClients() {
  const props = PropertiesService.getDocumentProperties();
  const clientsStr = props.getProperty('saved_clients');
  return clientsStr ? JSON.parse(clientsStr) : {};
}

/** Helper to save clients to properties */
function _saveClients(clients) {
  const props = PropertiesService.getDocumentProperties();
  props.setProperty('saved_clients', JSON.stringify(clients));
}

// Called from Setup.html to get saved clients
function getSavedClients() {
  return _getClients();
}

// Called from Setup.html to save a new client
function saveClient(nickname, details) {
  const clients = _getClients();
  clients[nickname] = details;
  _saveClients(clients);
  return clients;
}

function showManageClientsDialog() {
  const html = HtmlService.createHtmlOutputFromFile('ManageClients')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Manage Clients');
}

function deleteClient(nickname) {
  const clients = _getClients();
  if (clients[nickname]) {
    delete clients[nickname];
    _saveClients(clients);
  }
  return clients;
}

function renameClient(oldNickname, newNickname, details) {
  const clients = _getClients();
  if (oldNickname && oldNickname !== newNickname) {
    delete clients[oldNickname];
  }
  clients[newNickname] = details;
  _saveClients(clients);
  return clients;
}

function getWeekRangeLabel(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  
  // Get Monday
  const day = d.getDay();
  const diff = d.getDate() - day + (day === 0 ? -6 : 1);
  const monday = new Date(d.setDate(diff));
  
  // Get Sunday
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6);
  
  const options = { month: 'short', day: 'numeric' };
  return monday.toLocaleDateString('en-US', options) + ' - ' + sunday.toLocaleDateString('en-US', options);
}

// Helper to format currency
function formatCurrency(num) {
  return Number(num).toLocaleString('en-GB', {minimumFractionDigits: 2, maximumFractionDigits: 2});
}

function generateFromSelection() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const rangeList = sheet.getActiveRangeList();
    const ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];
    
    let data = [];
    let lowestRow = 0;
    
    for (let k = 0; k < ranges.length; k++) {
      const r = ranges[k];
      const rStart = r.getRow();
      const rNum = r.getNumRows();
      data = data.concat(sheet.getRange(rStart, 1, rNum, 4).getValues());
      
      const rEnd = rStart + rNum - 1;
      if (rEnd > lowestRow) lowestRow = rEnd;
    }
    
    let projectStr = "";
    // Find the first valid project string in the selection
    // Format: #0000 DESCRIPTION @CLIENT NICKNAME
    for (let i = 0; i < data.length; i++) {
       for (let j = 0; j < data[i].length; j++) {
         const cellVal = String(data[i][j]).trim();
         if (cellVal.match(/^#\S+\s+.*?\s+@/)) {
            projectStr = cellVal;
            break;
         }
       }
       if (projectStr) break;
    }
    
    const ui = SpreadsheetApp.getUi();
    
    if (!projectStr) {
       ui.alert("Could not find a project description in the highlighted cells.");
       return;
    }
    
    // Format: #0000 DESCRIPTION @CLIENT NICKNAME
    const match = projectStr.match(/#(\S+)\s+(.*?)\s+@(.*)/);
    if (!match) {
       ui.alert("Invalid project format.\nExpected: #0000 DESCRIPTION @CLIENT NICKNAME\nFound: " + projectStr);
       return;
    }
    
    const invoiceNum = match[1];
    const forText = match[2].trim();
    const clientNickname = match[3].trim();
    
    const clients = _getClients();
    const clientDetails = clients[clientNickname];
    if (!clientDetails) {
       ui.alert(`Client '${clientNickname}' not found. Please add them via 'Manage Clients...' first.`);
       return;
    }
    
    generateInvoiceHTML(clientDetails, invoiceNum, forText, clientNickname, data, lowestRow);
  } catch (error) {
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
  }
}

// Called from generateFromSelection to generate the actual invoice
function generateInvoiceHTML(clientDetails, invoiceNum, forText, clientNickname, data, lowestRow) {
  const weeksMap = {};
  let totalDays = 0;
  let totalFee = 0;
  const currency = "£"; // Currency symbol based on prompt example
  
  let minDate = null;
  let maxDate = null;
  
  let startIndex = 0;
  if (data.length > 0 && String(data[0][0]).trim().toUpperCase() === "DATE") {
    startIndex = 1;
  }
  
  for (let i = startIndex; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[0];
    const project = row[1];
    const type = row[2];
    const rateVal = row[3];
    
    // Check if type is "D" (case insensitive)
    if (String(type).trim().toUpperCase() !== "D") continue;
    
    // Parse date
    const d = new Date(dateVal);
    if (isNaN(d.getTime())) continue; // Skip invalid dates
    
    if (!minDate || d < minDate) minDate = new Date(d);
    if (!maxDate || d > maxDate) maxDate = new Date(d);
    
    const weekLabel = getWeekRangeLabel(d);
    
    // Parse rate
    let rate = 0;
    if (typeof rateVal === 'number') {
      rate = rateVal;
    } else if (typeof rateVal === 'string') {
      // extract numbers and decimals
      const matches = rateVal.match(/[0-9.]+/g);
      if (matches) {
        rate = parseFloat(matches.join(''));
      }
    }
    
    const groupKey = weekLabel + '_' + rate;
    
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
    const dateStr = ('0' + d.getDate()).slice(-2) + '/' + ('0' + (d.getMonth() + 1)).slice(-2);
    if (!weeksMap[groupKey].dates.includes(dateStr)) {
        weeksMap[groupKey].dates.push(dateStr);
    }

    let displayProject = project;
    if (project) {
      const m = String(project).match(/#(\S+)\s+(.*?)\s+@(.*)/);
      if (m) displayProject = m[2].trim();
    }
    
    if (displayProject) weeksMap[groupKey].projects.add(String(displayProject).trim());
    weeksMap[groupKey].total += rate;
    
    totalDays += 1;
    totalFee += rate;
  }
  
  // Format data for template
  const groupedData = [];
  for (const key in weeksMap) {
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
  groupedData.sort((a, b) => a.sortKey - b.sortKey);
  
  // Assign "Week X" labels
  if (groupedData.length > 0) {
    let currentWeekNum = 1;
    let currentWeekLabel = groupedData[0].originalWeekLabel;
    for (let j = 0; j < groupedData.length; j++) {
      if (groupedData[j].originalWeekLabel !== currentWeekLabel) {
        currentWeekNum++;
        currentWeekLabel = groupedData[j].originalWeekLabel;
      }
      groupedData[j].weekLabel = "Week " + currentWeekNum;
    }
  }
  
  let intervalStr = "";
  if (minDate && maxDate) {
    const options = { month: 'short', day: 'numeric', year: 'numeric' };
    if (minDate.getTime() === maxDate.getTime()) {
      intervalStr = minDate.toLocaleDateString('en-GB', options);
    } else {
      intervalStr = minDate.toLocaleDateString('en-GB', options) + ' - ' + maxDate.toLocaleDateString('en-GB', options);
    }
  } else {
    intervalStr = "No dates selected";
  }
  const documentTitle = `${invoiceNum} - Invoice Ric_${forText}`;
  
  const template = HtmlService.createTemplateFromFile('Invoice');
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
  
  let isUs = false;
  if (clientDetails) {
    const addrUpper = String(clientDetails).toUpperCase();
    const hasUs = /\b(US|USA|U\.S\.A\.|UNITED STATES|UNITED STATES OF AMERICA)\b/i.test(addrUpper);
    const hasUk = /\b(UK|U\.K\.|UNITED KINGDOM|GREAT BRITAIN|GB)\b/i.test(addrUpper);
    isUs = hasUs && !hasUk;
  }
  template.isUs = isUs;
  
  const now = new Date();
  const dateOptions = { day: 'numeric', month: 'long', year: 'numeric' };
  template.currentDate = now.toLocaleDateString('en-GB', dateOptions);
  
  const htmlOutput = template.evaluate()
    .setWidth(850)
    .setHeight(650)
    .setTitle(documentTitle);
    
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, 'Invoice');
}

function generateQuoteFromSelection() {
  try {
    const sheet = SpreadsheetApp.getActiveSheet();
    const rangeList = sheet.getActiveRangeList();
    const ranges = rangeList ? rangeList.getRanges() : [sheet.getActiveRange()];
    
    let data = [];
    let displayData = [];
    
    for (let k = 0; k < ranges.length; k++) {
      const r = ranges[k];
      const startRow = r.getRow();
      const numRows = r.getNumRows();
      data = data.concat(sheet.getRange(startRow, 1, numRows, 6).getValues());
      displayData = displayData.concat(sheet.getRange(startRow, 1, numRows, 6).getDisplayValues());
    }
    
    // Detect currency from display values
    let currencySymbol = "£"; // Default
    for (let r = 0; r < displayData.length; r++) {
      const rowText = displayData[r].join(" ");
      if (rowText.includes('$')) {
        currencySymbol = "$";
        break;
      }
      if (rowText.includes('£')) {
        currencySymbol = "£";
        break;
      }
    }
    
    let quoteDate = "";
    let projectDesc = "";
    let dayRate = "";
    const items = [];
    let totalAmount = "";
    let totalDays = 0;
    
    let headerFound = false;
    
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      
      const colA = String(row[0]).trim();
      const colB = String(row[1]).trim();
      const colC = String(row[2]).trim();
      const colD = String(row[3]).trim();
      const colF = String(row[5]).trim();
      
      // Parse top header section
      if (colA.toUpperCase() === "DATE") {
        const dateCell = row[1];
        if (Object.prototype.toString.call(dateCell) === '[object Date]') {
           quoteDate = dateCell.toDateString();
        } else {
           quoteDate = colB;
           const match = quoteDate.match(/(?:Quote\s*)?(\d+)\/(\d+)\/(\d+)/i);
           if (match) {
              const d = parseInt(match[1], 10);
              const m = parseInt(match[2], 10);
              let y = parseInt(match[3], 10);
              if (y < 100) y += 2000;
              const dateObj = new Date(y, m - 1, d);
              quoteDate = dateObj.toDateString(); 
           }
        }
      }
      if (colA.toUpperCase() === "PROJECT DESCRIPTION") {
        projectDesc = colB;
      }
      
      // Search for "DAY RATE" in any column
      let dayRateColIndex = -1;
      for (let col = 0; col < row.length; col++) {
         if (String(row[col]).trim().toUpperCase() === "DAY RATE") {
            dayRateColIndex = col;
            break;
         }
      }
      if (dayRateColIndex !== -1) {
         if (i + 1 < data.length) {
            const rawDayRate = String(data[i+1][dayRateColIndex]).trim();
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
           const dVal = parseFloat(colD);
           if (!isNaN(dVal)) totalDays += dVal;
        }
      }
    }
    
    const ui = SpreadsheetApp.getUi();
    
    if (!quoteDate && !projectDesc && items.length === 0) {
       ui.alert("Could not find quote data. Please highlight the entire quote block (from DATE to Total).");
       return;
    }
    
    function customFormatCurrency(val) {
      const str = String(val).trim();
      if (!str) return "";
      
      const matches = str.match(/[0-9.,]+/g);
      if (matches) {
         const numStr = matches.join('').replace(/,/g, '');
         const num = parseFloat(numStr);
         if (!isNaN(num)) {
            return num.toLocaleString('en-GB', {minimumFractionDigits: 0, maximumFractionDigits: 2});
         }
      }
      return str;
    }
    
    const template = HtmlService.createTemplateFromFile('Quote');
    template.documentTitle = "Quote - " + projectDesc;
    template.quoteDate = quoteDate;
    template.clientDetails = projectDesc; // Using project description for "For:" field
    template.dayRate = customFormatCurrency(dayRate);
    
    // Format all item amounts
    for (let j = 0; j < items.length; j++) {
       items[j].totalDisplay = customFormatCurrency(items[j].totalDisplay);
    }
    
    template.groupedData = items;
    template.totalDays = totalDays;
    template.totalFeeDisplay = customFormatCurrency(totalAmount);
    template.currency = currencySymbol;
    
    const htmlOutput = template.evaluate()
      .setWidth(850)
      .setHeight(650)
      .setTitle(template.documentTitle);
      
    ui.showModalDialog(htmlOutput, 'Quote');
  } catch (error) {
    SpreadsheetApp.getUi().alert("An error occurred: " + error.message);
  }
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
 * @param {string} forText
 * @return {string} The URL of the saved file.
 */
function savePdfToDrive(base64Data, fileName, invoiceNum, lowestRow, sendEmail, forText) {
  try {
    const folder = DriveApp.getFolderById(CONFIG.DRIVE_FOLDER_ID);
    const contentType = "application/pdf";
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, contentType, fileName);
    const file = folder.createFile(blob);
    
    // Auto-tag the sheet with the invoice number as a hyperlink to the PDF
    if (invoiceNum && lowestRow) {
      const sheet = SpreadsheetApp.getActiveSheet();
      const fileUrl = file.getUrl();
      
      const cleanInvoiceNum = String(invoiceNum).replace(/["']/g, "");
      const richText = SpreadsheetApp.newRichTextValue()
        .setText("#" + cleanInvoiceNum)
        .setLinkUrl(fileUrl)
        .build();
        
      sheet.getRange(lowestRow, 7).setRichTextValue(richText);
      sheet.getRange(lowestRow, 8).setValue(new Date()); // Column H is the generation date
      sheet.getRange(lowestRow, 9).setValue("N"); // Column I is "N"
    }
    
    // Send email with PDF attachment if requested
    if (sendEmail) {
      const cleanInvoiceNum = String(invoiceNum).replace(/["']/g, "");
      const cleanForText = String(forText || "").replace(/["']/g, "").trim();
      MailApp.sendEmail({
        to: CONFIG.DEFAULT_EMAIL_RECIPIENTS,
        subject: `Ric invoice #${cleanInvoiceNum}`,
        body: `Heya,\n\nPlease find attached invoice #${cleanInvoiceNum} for ${cleanForText}\n\nBest,\nRic`,
        attachments: [blob]
      });
    }
    
    return file.getUrl();
  } catch (e) {
    throw new Error("Failed to save PDF to Google Drive: " + e.toString());
  }
}
