/**
 * Expense Tracker & Approval System
 * Google Apps Script Web Application
 */

// ============== CONFIGURATION ==============
const CONFIG = {
  SHEET_NAMES: {
    USERS: 'Users',
    EXPENSE_CLAIMS: 'ExpenseClaims',
    DEPARTMENT_BUDGETS: 'DepartmentBudgets',
    AUDIT_LOG: 'AuditLog'
  },
  CATEGORIES: ['Travel', 'Meals', 'Supplies', 'Equipment', 'Training', 'Communication', 'Other'],
  STATUS: {
    DRAFT: 'Draft',
    PENDING_MANAGER: 'Pending Manager',
    PENDING_FINANCE: 'Pending Finance',
    APPROVED: 'Approved',
    REJECTED: 'Rejected',
    PAID: 'Paid'
  },
  THRESHOLD: 5000,
  ROLES: {
    EMPLOYEE: 'Employee',
    MANAGER: 'Manager',
    FINANCE: 'Finance Admin'
  }
};


const SHEET_CONFIG = {
  Users: { columns: ['Email', 'Name', 'Role', 'Department', 'Active'], sampleData: [['admin@company.com', 'Admin User', 'Admin', 'Finance', true]] },
  ExpenseClaims: { columns: ['ClaimID', 'EmployeeEmail', 'EmployeeName', 'Category', 'Amount', 'ExpenseDate', 'Description', 'ReceiptLink', 'Status', 'ManagerEmail', 'ManagerApprovedAt', 'ManagerComments', 'FinanceApprovedAt', 'FinanceComments', 'PaidAt', 'SubmittedAt', 'Department'] },
  DepartmentBudgets: { columns: ['Department', 'Month', 'Year', 'Allocated', 'Spent', 'Remaining'] },
  AuditLog: { columns: ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue'] }
};


// ============== ENHANCED UTILITIES (v2.0) ==============
const VERSION = '2.0.0';

/**
 * Initialize all required sheets for this application
 */
function initializeSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const created = [];
  
  for (const [sheetName, config] of Object.entries(SHEET_CONFIG)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      created.push(sheetName);
      sheet.getRange(1, 1, 1, config.columns.length).setValues([config.columns]).setFontWeight('bold');
      sheet.setFrozenRows(1);
      if (config.sampleData) config.sampleData.forEach(row => sheet.appendRow(row));
    }
  }
  Logger.log('InitializeSheets: Created ' + created.join(', '));
  return { created };
}

function handleError(error, context) {
  const errorMsg = error instanceof Error ? error.message : String(error);
  Logger.log('[ERROR] ' + context + ': ' + errorMsg);
  if (typeof logAction === 'function') logAction('ERROR', context, '', errorMsg);
  return { success: false, error: errorMsg };
}

function backupData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const backupName = 'Backup_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  ss.copy(backupName);
  if (typeof logAction === 'function') logAction('BACKUP_CREATED', 'System', '', backupName);
  return { success: true, backupName };
}

function exportToPDF(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getActiveSheet();
  const pdf = DriveApp.getFileById(ss.getId()).getAs('application/pdf');
  const pdfName = sheet.getName() + '_' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd') + '.pdf';
  DriveApp.getRootFolder().createFile(pdf).setName(pdfName);
  return { success: true, fileName: pdfName };
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎯 System Menu')
    .addItem('📊 Initialize Sheets', 'initializeSheets')
    .addItem('💾 Create Backup', 'backupData')
    .addItem('📄 Export to PDF', 'exportToPDF')
    .addSeparator()
    .addItem('ℹ️ About', 'showAbout')
    .addToUi();
}

function showAbout() {
  const ui = SpreadsheetApp.getUi();
  ui.alert('Expense Tracker v' + VERSION, 'Enhanced with:\n- initializeSheets()\n- backupData()\n- exportToPDF()', ui.ButtonSet.OK);
}


// ============== AUTHENTICATION ==============
function getUserRole(email) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase() && data[i][4] === true) {
      return data[i][2];
    }
  }
  return null;
}

function getCurrentUser() {
  const email = Session.getActiveUser().getEmail();
  const user = getUserByEmail(email);
  return user || { email, name: 'Unknown', role: null, department: null };
}

function getUserByEmail(email) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0].toLowerCase() === email.toLowerCase()) {
      return {
        email: data[i][0],
        name: data[i][1],
        role: data[i][2],
        department: data[i][3],
        active: data[i][4]
      };
    }
  }
  return null;
}

function requireRole(allowedRoles) {
  const user = getCurrentUser();
  if (!user.role || !allowedRoles.includes(user.role)) {
    throw new Error('Unauthorized: Insufficient permissions');
  }
  return user;
}

// ============== SHEET HELPERS ==============
function getSheet(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    setupHeaders(sheet, name);
  }
  return sheet;
}

function setupHeaders(sheet, name) {
  const headers = {
    'Users': ['Email', 'Name', 'Role', 'Department', 'Active'],
    'ExpenseClaims': ['ClaimID', 'EmployeeEmail', 'EmployeeName', 'Category', 'Amount', 
                     'ExpenseDate', 'Description', 'ReceiptLink', 'Status', 'ManagerEmail',
                     'ManagerApprovedAt', 'ManagerComments', 'FinanceApprovedAt', 'FinanceComments',
                     'PaidAt', 'SubmittedAt', 'Department'],
    'DepartmentBudgets': ['Department', 'Month', 'Year', 'Allocated', 'Spent', 'Remaining'],
    'AuditLog': ['Timestamp', 'User', 'Action', 'RecordID', 'OldValue', 'NewValue']
  };
  if (headers[name]) {
    sheet.getRange(1, 1, 1, headers[name].length).setValues([headers[name]]);
    sheet.getRange(1, 1, 1, headers[name].length).setFontWeight('bold');
  }
}

function generateId() {
  return 'EXP-' + Utilities.getUuid().substring(0, 8).toUpperCase();
}

function logAction(action, recordId, oldValue, newValue) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.AUDIT_LOG);
  sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, recordId, oldValue, newValue]);
}

// ============== NOTIFICATIONS ==============
function sendNotification(toEmail, subject, htmlBody) {
  try {
    MailApp.sendEmail({ to: toEmail, subject, htmlBody, name: 'Expense Tracker' });
    return true;
  } catch (e) {
    console.error('Email failed:', e);
    return false;
  }
}

function getEmailTemplate(status, employeeName, category, amount, description) {
  const colors = {
    'Pending Manager': '#F59E0B',
    'Pending Finance': '#8B5CF6',
    'Approved': '#10B981',
    'Rejected': '#EF4444',
    'Paid': '#059669'
  };
  return `
    <div style="font-family: -apple-system, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px;">
      <div style="background: linear-gradient(135deg, ${colors[status] || '#6B7280'} 0%, ${colors[status] ? adjustColor(colors[status]) : '#6B7280'} 100%); padding: 30px; border-radius: 16px 16px 0 0;">
        <h1 style="color: white; margin: 0;">Expense Claim ${status}</h1>
      </div>
      <div style="background: white; padding: 30px; border-radius: 0 0 16px 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.1);">
        <p style="color: #374151;">Dear ${employeeName},</p>
        <p style="color: #6B7280;">Your expense claim has been <strong style="color: ${colors[status] || '#6B7280'};">${status}</strong></p>
        <div style="background: #F9FAFB; padding: 20px; border-radius: 12px; margin: 20px 0;">
          <p style="margin: 0 0 10px;"><strong>Category:</strong> ${category}</p>
          <p style="margin: 0 0 10px;"><strong>Amount:</strong> ₹${amount.toLocaleString()}</p>
          <p style="margin: 0;"><strong>Description:</strong> ${description}</p>
        </div>
        <div style="text-align: center; margin-top: 30px;">
          <a href="${ScriptApp.getService().getUrl()}" style="background: ${colors[status] || '#6B7280'}; color: white; padding: 12px 30px; border-radius: 25px; text-decoration: none; font-weight: 600;">View Details</a>
        </div>
      </div>
    </div>
  `;
}

function adjustColor(hex) {
  return hex;
}

function getApprovalEmailTemplate(claimId, employeeName, category, amount, description, pendingRole) {
  const approveColor = '#10B981';
  const rejectColor = '#EF4444';
  return `
    <div style="font-family: -apple-system, sans-serif; max-width: 600px; margin: 0 auto;">
      <div style="background: linear-gradient(135deg, #667EEA 0%, #764BA2 100%); padding: 30px; border-radius: 16px 16px 0 0;">
        <h1 style="color: white; margin: 0;">Expense Claim Approval Required</h1>
      </div>
      <div style="background: white; padding: 30px; border-radius: 0 0 16px 16px; box-shadow: 0 4px 20px rgba(0,0,0,0.1);">
        <p style="color: #374151; font-size: 16px;">Hi ${pendingRole},</p>
        <p style="color: #6B7280;"><strong>${employeeName}</strong> has submitted an expense claim requiring your approval.</p>
        <div style="background: #F9FAFB; padding: 20px; border-radius: 12px; margin: 20px 0;">
          <p style="margin: 0 0 10px;"><strong>Category:</strong> ${category}</p>
          <p style="margin: 0 0 10px;"><strong>Amount:</strong> ₹${amount.toLocaleString()}</p>
          <p style="margin: 0;"><strong>Description:</strong> ${description}</p>
        </div>
        <div style="display: flex; gap: 12px; margin-top: 24px;">
          <a href="${ScriptApp.getService().getUrl()}?action=view&claimId=${claimId}" style="flex: 1; background: ${approveColor}; color: white; padding: 14px; border-radius: 12px; text-decoration: none; font-weight: 600; text-align: center;">Review in System</a>
        </div>
      </div>
    </div>
  `;
}

// ============== BUDGET CHECK ==============
function checkDepartmentBudget(department, amount, month, year) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.DEPARTMENT_BUDGETS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === department && data[i][1] === month && data[i][2] == year) {
      const remaining = data[i][5];
      return {
        hasBudget: remaining >= amount,
        allocated: data[i][3],
        spent: data[i][4],
        remaining: remaining,
        requiresFinanceSignoff: remaining < amount
      };
    }
  }
  return { hasBudget: false, requiresFinanceSignoff: true, remaining: 0 };
}

function updateDepartmentSpend(department, month, year, amount) {
  const sheet = getSheet(CONFIG.SHEET_NAMES.DEPARTMENT_BUDGETS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === department && data[i][1] === month && data[i][2] == year) {
      const row = i + 1;
      const newSpent = data[i][4] + amount;
      const newRemaining = data[i][3] - newSpent;
      sheet.getRange(row, 5, 1, 1).setValue(newSpent);
      sheet.getRange(row, 6, 1, 1).setValue(Math.max(0, newRemaining));
      return;
    }
  }
}

function getManagerForEmployee(email) {
  const user = getUserByEmail(email);
  if (!user) return null;
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === CONFIG.ROLES.MANAGER && data[i][3] === user.department) {
      return data[i][0];
    }
  }
  return null;
}

function getFinanceAdmin() {
  const sheet = getSheet(CONFIG.SHEET_NAMES.USERS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][2] === CONFIG.ROLES.FINANCE) {
      return data[i][0];
    }
  }
  return null;
}

// ============== API FUNCTIONS ==============
function submitExpenseClaim(category, amount, expenseDate, description, receiptLink) {
  const user = requireRole([CONFIG.ROLES.EMPLOYEE, CONFIG.ROLES.MANAGER, CONFIG.ROLES.FINANCE]);
  
  if (!CONFIG.CATEGORIES.includes(category)) {
    throw new Error('Invalid category');
  }
  
  if (amount <= 0) {
    throw new Error('Amount must be positive');
  }
  
  const date = new Date(expenseDate);
  const month = date.toLocaleString('en-US', { month: 'long' });
  const year = date.getFullYear();
  
  // Budget check
  const budgetCheck = checkDepartmentBudget(user.department, amount, month, year);
  
  const claimId = generateId();
  const managerEmail = getManagerForEmployee(user.email);
  const now = new Date();
  
  let status = CONFIG.STATUS.PENDING_MANAGER;
  if (amount > CONFIG.THRESHOLD) {
    status = CONFIG.STATUS.PENDING_MANAGER;
  }
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  sheet.appendRow([
    claimId, user.email, user.name, category, amount, new Date(expenseDate),
    description, receiptLink || '', status, managerEmail, '', '', '', '',
    '', now, user.department
  ]);
  
  logAction('SUBMIT_EXPENSE', claimId, '', status);
  
  // Notify manager
  if (managerEmail) {
    sendNotification(managerEmail, `Expense Claim from ${user.name}`, 
      getApprovalEmailTemplate(claimId, user.name, category, amount, description, 'Manager'));
  }
  
  return { 
    success: true, 
    claimId,
    requiresFinanceSignoff: budgetCheck.requiresFinanceSignoff
  };
}

function approveByManager(claimId, comments) {
  requireRole([CONFIG.ROLES.MANAGER, CONFIG.ROLES.FINANCE]);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === claimId) {
      const row = i + 1;
      const oldStatus = data[i][8];
      const amount = data[i][4];
      const userEmail = data[i][1];
      
      if (oldStatus !== CONFIG.STATUS.PENDING_MANAGER) {
        throw new Error('Claim is not pending manager approval');
      }
      
      // Check if needs finance approval
      let newStatus = CONFIG.STATUS.APPROVED;
      if (amount > CONFIG.THRESHOLD) {
        newStatus = CONFIG.STATUS.PENDING_FINANCE;
        const financeEmail = getFinanceAdmin();
        if (financeEmail) {
          sendNotification(financeEmail, `Expense Claim - Finance Approval Required`,
            getApprovalEmailTemplate(claimId, data[i][2], data[i][3], amount, data[i][6], 'Finance Admin'));
        }
      } else {
        // Update budget
        const date = new Date(data[i][5]);
        updateDepartmentSpend(data[i][16], date.toLocaleString('en-US', { month: 'long' }), date.getFullYear(), amount);
      }
      
      sheet.getRange(row, 10, 1, 1).setValue(new Date());
      sheet.getRange(row, 11, 1, 1).setValue(comments || '');
      sheet.getRange(row, 8, 1, 1).setValue(newStatus);
      
      logAction('MANAGER_APPROVE', claimId, oldStatus, newStatus);
      
      // Notify employee
      sendNotification(userEmail, 'Expense Claim Update',
        getEmailTemplate(newStatus, data[i][2], data[i][3], amount, data[i][6]));
      
      return { success: true, newStatus };
    }
  }
  
  throw new Error('Claim not found');
}

function approveByFinance(claimId, comments) {
  requireRole([CONFIG.ROLES.FINANCE]);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === claimId) {
      const row = i + 1;
      const oldStatus = data[i][8];
      
      if (oldStatus !== CONFIG.STATUS.PENDING_FINANCE) {
        throw new Error('Claim is not pending finance approval');
      }
      
      const amount = data[i][4];
      const userEmail = data[i][1];
      
      // Update budget
      const date = new Date(data[i][5]);
      updateDepartmentSpend(data[i][16], date.toLocaleString('en-US', { month: 'long' }), date.getFullYear(), amount);
      
      sheet.getRange(row, 12, 1, 1).setValue(new Date());
      sheet.getRange(row, 13, 1, 1).setValue(comments || '');
      sheet.getRange(row, 8, 1, 1).setValue(CONFIG.STATUS.APPROVED);
      
      logAction('FINANCE_APPROVE', claimId, oldStatus, CONFIG.STATUS.APPROVED);
      
      // Notify employee
      sendNotification(userEmail, 'Expense Claim Approved',
        getEmailTemplate(CONFIG.STATUS.APPROVED, data[i][2], data[i][3], amount, data[i][6]));
      
      return { success: true };
    }
  }
  
  throw new Error('Claim not found');
}

function rejectClaim(claimId, reason) {
  if (!reason) throw new Error('Rejection reason is required');
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === claimId) {
      const row = i + 1;
      const oldStatus = data[i][8];
      const userEmail = data[i][1];
      
      sheet.getRange(row, 8, 1, 1).setValue(CONFIG.STATUS.REJECTED);
      
      // Store rejection reason in comments based on stage
      if (oldStatus === CONFIG.STATUS.PENDING_MANAGER) {
        sheet.getRange(row, 11, 1, 1).setValue(reason);
      } else {
        sheet.getRange(row, 13, 1, 1).setValue(reason);
      }
      
      logAction('REJECT_CLAIM', claimId, oldStatus, CONFIG.STATUS.REJECTED);
      
      // Notify employee
      sendNotification(userEmail, 'Expense Claim Rejected',
        getEmailTemplate(CONFIG.STATUS.REJECTED, data[i][2], data[i][3], data[i][4], data[i][6]));
      
      return { success: true };
    }
  }
  
  throw new Error('Claim not found');
}

function markAsPaid(claimId) {
  requireRole([CONFIG.ROLES.FINANCE]);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === claimId) {
      const row = i + 1;
      const oldStatus = data[i][8];
      
      if (oldStatus !== CONFIG.STATUS.APPROVED) {
        throw new Error('Only approved claims can be marked as paid');
      }
      
      sheet.getRange(row, 14, 1, 1).setValue(new Date());
      sheet.getRange(row, 8, 1, 1).setValue(CONFIG.STATUS.PAID);
      
      logAction('MARK_PAID', claimId, oldStatus, CONFIG.STATUS.PAID);
      
      // Notify employee
      sendNotification(data[i][1], 'Expense Claim Paid',
        getEmailTemplate(CONFIG.STATUS.PAID, data[i][2], data[i][3], data[i][4], data[i][6]));
      
      return { success: true };
    }
  }
  
  throw new Error('Claim not found');
}

// ============== DATA RETRIEVAL ==============
function getMyClaims() {
  const user = getCurrentUser();
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  const claims = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][1].toLowerCase() === user.email.toLowerCase()) {
      claims.push({
        claimId: data[i][0],
        category: data[i][3],
        amount: data[i][4],
        expenseDate: data[i][5],
        description: data[i][6],
        receiptLink: data[i][7],
        status: data[i][8],
        submittedAt: data[i][15]
      });
    }
  }
  
  return claims;
}

function getPendingManagerApprovals() {
  const user = getCurrentUser();
  if (user.role !== CONFIG.ROLES.MANAGER && user.role !== CONFIG.ROLES.FINANCE) {
    throw new Error('Unauthorized');
  }
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  const claims = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === CONFIG.STATUS.PENDING_MANAGER) {
      // Check if this is for their department or they're finance
      if (user.role === CONFIG.ROLES.FINANCE || data[i][9] === user.email) {
        claims.push({
          claimId: data[i][0],
          employeeEmail: data[i][1],
          employeeName: data[i][2],
          category: data[i][3],
          amount: data[i][4],
          expenseDate: data[i][5],
          description: data[i][6],
          receiptLink: data[i][7],
          submittedAt: data[i][15]
        });
      }
    }
  }
  
  return claims;
}

function getPendingFinanceApprovals() {
  requireRole([CONFIG.ROLES.FINANCE]);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  const claims = [];
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][8] === CONFIG.STATUS.PENDING_FINANCE) {
      claims.push({
        claimId: data[i][0],
        employeeEmail: data[i][1],
        employeeName: data[i][2],
        category: data[i][3],
        amount: data[i][4],
        expenseDate: data[i][5],
        description: data[i][6],
        receiptLink: data[i][7],
        submittedAt: data[i][15]
      });
    }
  }
  
  return claims;
}

function getDashboardData() {
  const user = getCurrentUser();
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  
  let myTotal = 0, approvedTotal = 0, pendingTotal = 0;
  const categoryBreakdown = {};
  const monthlySpend = {};
  
  for (let i = 1; i < data.length; i++) {
    const isMyClaim = data[i][1].toLowerCase() === user.email.toLowerCase();
    const amount = data[i][4];
    const category = data[i][3];
    const status = data[i][8];
    const month = new Date(data[i][5]).toLocaleString('en-US', { month: 'short' });
    
    if (isMyClaim) {
      myTotal += amount;
      if (status === CONFIG.STATUS.APPROVED || status === CONFIG.STATUS.PAID) {
        approvedTotal += amount;
      } else if (status !== CONFIG.STATUS.REJECTED) {
        pendingTotal += amount;
      }
    }
    
    if (status === CONFIG.STATUS.APPROVED || status === CONFIG.STATUS.PAID) {
      categoryBreakdown[category] = (categoryBreakdown[category] || 0) + amount;
      monthlySpend[month] = (monthlySpend[month] || 0) + amount;
    }
  }
  
  return { myTotal, approvedTotal, pendingTotal, categoryBreakdown, monthlySpend };
}

function getAllClaims() {
  requireRole([CONFIG.ROLES.FINANCE]);
  
  const sheet = getSheet(CONFIG.SHEET_NAMES.EXPENSE_CLAIMS);
  const data = sheet.getDataRange().getValues();
  const claims = [];
  
  for (let i = 1; i < data.length; i++) {
    claims.push({
      claimId: data[i][0],
      employeeEmail: data[i][1],
      employeeName: data[i][2],
      category: data[i][3],
      amount: data[i][4],
      expenseDate: data[i][5],
      description: data[i][6],
      receiptLink: data[i][7],
      status: data[i][8],
      managerEmail: data[i][9],
      managerApprovedAt: data[i][10],
      managerComments: data[i][11],
      financeApprovedAt: data[i][12],
      financeComments: data[i][13],
      paidAt: data[i][14],
      submittedAt: data[i][15],
      department: data[i][16]
    });
  }
  
  return claims;
}

// ============== WEB APP ==============
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Expense Tracker')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}
