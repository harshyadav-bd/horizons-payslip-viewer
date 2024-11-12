// Global variables
const SHEET_ID = '1ckvUpn3Nnbe6qMUt0VM9pabYEX1jFx0aVDzlZoBJ1nk';
const SEARCH_QUERY = 'subject:"Horizons - Payslip"';

function doGet() {
  return HtmlService.createHtmlOutput(createWebAppHTML())
    .setTitle('Payslip Viewer')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function scanEmails() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    // Clear existing data
    if (sheet.getLastRow() > 1) {
      sheet.getRange('A2:B' + sheet.getLastRow()).clear();
    }

    let row = 2;
    let processedCount = 0;
    const threads = GmailApp.search(SEARCH_QUERY);

    if (!threads || threads.length === 0) {
      return 'No emails found matching the search criteria.';
    }

    threads.forEach(thread => {
      // Get all messages in the thread instead of just the first one
      const messages = thread.getMessages();
      
      // Process each message in the thread
      messages.forEach(message => {
        const subject = message.getSubject();
        const body = message.getPlainBody();
        const monthMatch = subject.match(/Horizons - Payslip for (\w+ \d{4})/);
        const nameMatch = body.match(/Hi (\w+)/);

        if (monthMatch && nameMatch) {
          // Store exact matched values
          sheet.getRange(row, 1).setValue(nameMatch[1].trim());
          sheet.getRange(row, 2).setValue(monthMatch[1].trim());
          row++;
          processedCount++;
        }
      });
    });

    return `Scan complete. Processed ${processedCount} emails.`;
  } catch (error) {
    Logger.log('Error in scanEmails: ' + error.toString());
    return 'Error during scan: ' + error.toString();
  }
}

function getMonths() {
  try {
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    if (sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange('B2:B' + sheet.getLastRow()).getValues();
    const months = data.flat()
                     .filter(item => item && item.toString().trim() !== '')
                     .map(item => item.toString().trim());
    
    const uniqueMonths = [...new Set(months)];
    
    // Sort months by date
    return uniqueMonths.sort((a, b) => {
      const dateA = new Date(a);
      const dateB = new Date(b);
      return dateB - dateA;  // Newest first
    });
  } catch (error) {
    Logger.log('Error in getMonths: ' + error.toString());
    return [];
  }
}

function getEmployeesByMonth(month) {
  try {
    if (!month) return [];
    const sheet = SpreadsheetApp.openById(SHEET_ID).getActiveSheet();
    if (sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
    
    // Use Set to get unique employee names
    const uniqueEmployees = new Set(
      data
        .filter(row => row[1].toString().trim() === month.toString().trim())
        .map(row => row[0].toString().trim())
        .filter(name => name !== '')
    );
    
    // Convert Set back to sorted array
    const employees = Array.from(uniqueEmployees).sort();
    
    Logger.log('Month requested: ' + month);
    Logger.log('Unique employees found: ' + JSON.stringify(employees));
    
    return employees;
  } catch (error) {
    Logger.log('Error in getEmployeesByMonth: ' + error.toString());
    return [];
  }
}

function createWebAppHTML() {
  return `<!DOCTYPE html>
<html>
<head>
  <base target="_top">
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <style>
    body {
      font-family: Arial, sans-serif;
      line-height: 1.6;
      margin: 0;
      padding: 20px;
      background-color: #f5f5f5;
    }
    .container {
      max-width: 800px;
      margin: 0 auto;
      background-color: white;
      padding: 20px;
      border-radius: 8px;
      box-shadow: 0 2px 4px rgba(0,0,0,0.1);
    }
    h2 {
      color: #333;
      margin-bottom: 20px;
      text-align: center;
    }
    .control-group {
      margin-bottom: 20px;
    }
    select, button {
      width: 100%;
      padding: 10px;
      margin: 5px 0;
      border: 1px solid #ddd;
      border-radius: 4px;
      font-size: 14px;
    }
    button {
      background-color: #4CAF50;
      color: white;
      border: none;
      cursor: pointer;
      font-weight: bold;
      transition: background-color 0.3s;
    }
    button:hover {
      background-color: #45a049;
    }
    button:disabled {
      background-color: #cccccc;
      cursor: not-allowed;
    }
    .status {
      padding: 10px;
      margin: 10px 0;
      border-radius: 4px;
      display: none;
    }
    .success {
      background-color: #dff0d8;
      color: #3c763d;
      border: 1px solid #d6e9c6;
    }
    .error {
      background-color: #f2dede;
      color: #a94442;
      border: 1px solid #ebccd1;
    }
    .loading {
      display: none;
      text-align: center;
      color: #666;
      font-style: italic;
      margin: 10px 0;
    }
    #employeeList {
      margin-top: 20px;
    }
    #employeeList h3 {
      color: #333;
      border-bottom: 2px solid #eee;
      padding-bottom: 10px;
    }
    ul {
      list-style-type: none;
      padding: 0;
    }
    li {
      padding: 10px;
      border-bottom: 1px solid #eee;
      transition: background-color 0.2s;
    }
    li:hover {
      background-color: #f9f9f9;
    }
    li:last-child {
      border-bottom: none;
    }
  </style>
</head>
<body>
  <div class="container">
    <h2>Horizons Payslip Viewer</h2>
    <div class="control-group">
      <button id="scanButton" onclick="handleScan()">Scan Emails</button>
      <div id="scanStatus" class="status"></div>
      <div id="loading" class="loading">Scanning emails, please wait...</div>
    </div>
    <div class="control-group">
      <select id="monthSelect" onchange="handleMonthChange()">
        <option value="">Select Month</option>
      </select>
    </div>
    <div id="employeeList"></div>
  </div>

  <script>
    window.onload = function() {
      loadMonths();
    };

    function loadMonths() {
      const select = document.getElementById('monthSelect');
      select.disabled = true;
      google.script.run
        .withSuccessHandler(function(months) {
          select.innerHTML = '<option value="">Select Month</option>';
          if (months && months.length > 0) {
            months.forEach(function(month) {
              if (month) {
                const option = document.createElement('option');
                option.value = month;
                option.textContent = month;
                select.appendChild(option);
              }
            });
          }
          select.disabled = false;
        })
        .withFailureHandler(function(error) {
          showStatus('Failed to load months: ' + error, false);
          select.disabled = false;
        })
        .getMonths();
    }

    function handleMonthChange() {
      const month = document.getElementById('monthSelect').value;
      const employeeList = document.getElementById('employeeList');
      
      if (!month) {
        employeeList.innerHTML = '';
        return;
      }

      google.script.run
        .withSuccessHandler(function(employees) {
          if (!employees || employees.length === 0) {
            employeeList.innerHTML = '<h3>No employees found for ' + month + '</h3>';
            return;
          }
          let html = '<h3>Employees for ' + month + '</h3><ul>';
          employees.forEach(function(employee) {
            if (employee) {
              html += '<li>' + employee + '</li>';
            }
          });
          html += '</ul>';
          employeeList.innerHTML = html;
        })
        .withFailureHandler(function(error) {
          showStatus('Failed to load employees: ' + error, false);
        })
        .getEmployeesByMonth(month);
    }

    function handleScan() {
      const button = document.getElementById('scanButton');
      const loading = document.getElementById('loading');
      button.disabled = true;
      loading.style.display = 'block';
      
      google.script.run
        .withSuccessHandler(function(result) {
          showStatus(result, true);
          button.disabled = false;
          loading.style.display = 'none';
          loadMonths();
        })
        .withFailureHandler(function(error) {
          showStatus('Scan failed: ' + error, false);
          button.disabled = false;
          loading.style.display = 'none';
        })
        .scanEmails();
    }

    function showStatus(message, isSuccess) {
      const status = document.getElementById('scanStatus');
      status.textContent = message;
      status.className = 'status ' + (isSuccess ? 'success' : 'error');
      status.style.display = 'block';
      setTimeout(function() {
        status.style.display = 'none';
      }, 5000);
    }
  </script>
</body>
</html>`;
}
