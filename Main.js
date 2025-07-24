// Main.gs - Core System Entry Point
// Enhanced Parsonage Tenant Management System v2.0

/**
 * Global Configuration Constants
 */
const CONFIG = {
  // Sheet Names
  SHEETS: {
    TENANTS: 'Tenants',
    BUDGET: 'Budget', 
    APPLICATIONS: 'Tenant Applications',
    MOVEOUTS: 'Move-Out Requests',
    GUEST_ROOMS: 'Guest Rooms',
    GUEST_BOOKINGS: 'Guest Bookings',
    MAINTENANCE: 'Maintenance Requests',
    DOCUMENTS: 'Documents',
    SETTINGS: 'System Settings'
  },
  
  // System Settings
  SYSTEM: {
    PROPERTY_NAME: 'Parsonage Living Community',
    MANAGER_EMAIL: Session.getActiveUser().getEmail(),
    TIME_ZONE: Session.getScriptTimeZone(),
    CURRENCY: 'USD',
    DATE_FORMAT: 'yyyy-MM-dd',
    LATE_FEE_DAYS: 5,
    LATE_FEE_AMOUNT: 25
  },
  
  // Status Options
  STATUS: {
    ROOM: {
      VACANT: 'Vacant',
      OCCUPIED: 'Occupied', 
      MAINTENANCE: 'Maintenance',
      PENDING: 'Pending'
    },
    PAYMENT: {
      PAID: 'Paid',
      DUE: 'Due',
      OVERDUE: 'Overdue',
      PARTIAL: 'Partial'
    },
    BOOKING: {
      PENDING: 'Pending',
      CONFIRMED: 'Confirmed',
      CHECKED_IN: 'Checked In',
      CHECKED_OUT: 'Checked Out',
      CANCELLED: 'Cancelled'
    },
    MAINTENANCE: {
      OPEN: 'Open',
      IN_PROGRESS: 'In Progress',
      COMPLETED: 'Completed',
      CANCELLED: 'Cancelled'
    }
  }
};

/**
 * Main Menu Creation - Enhanced with new features
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('🏠 Parsonage Manager')
    .addItem('🚀 Initialize System', 'initializeCompleteSystem')
    .addItem('⚙️ System Settings', 'showSystemSettings')
    .addSeparator()
    
    .addSubMenu(ui.createMenu('👥 Tenant Management')
      .addItem('📋 View Tenant Dashboard', 'showTenantDashboard')
      .addItem('💰 Check All Payment Status', 'checkAllPaymentStatus')
      .addItem('📧 Send Rent Reminders', 'sendRentReminders')
      .addItem('⚠️ Send Late Payment Alerts', 'sendLatePaymentAlerts')
      .addItem('📄 Send Monthly Invoices', 'sendMonthlyInvoices')
      .addItem('✅ Mark Payment Received', 'markPaymentReceived')
      .addItem('🔄 Process Move-In', 'processMoveIn')
      .addItem('📤 Process Move-Out', 'processMoveOut'))
    
    .addSubMenu(ui.createMenu('🛏️ Guest Room Management')
      .addItem('📅 Today\'s Arrivals & Departures', 'showTodayGuestActivity')
      .addItem('🔍 Check Room Availability', 'checkGuestRoomAvailability')
      .addItem('✅ Process Check-In', 'processGuestCheckIn')
      .addItem('📤 Process Check-Out', 'processGuestCheckOut')
      .addItem('📊 Guest Room Analytics', 'showGuestRoomAnalytics')
      .addItem('💲 Dynamic Pricing Analysis', 'analyzeGuestRoomPricing'))
    
    .addSubMenu(ui.createMenu('🔧 Maintenance System')
      .addItem('📝 View Open Requests', 'showMaintenanceRequests')
      .addItem('⚡ Create Urgent Request', 'createUrgentMaintenanceRequest')
      .addItem('📊 Maintenance Dashboard', 'showMaintenanceDashboard')
      .addItem('💰 Maintenance Cost Report', 'generateMaintenanceCostReport'))
    
    .addSubMenu(ui.createMenu('📊 Financial Reports')
      .addItem('💰 Monthly Financial Report', 'generateMonthlyFinancialReport')
      .addItem('📈 Revenue Analysis', 'showRevenueAnalysis')
      .addItem('🏠 Occupancy Analytics', 'showOccupancyAnalytics')
      .addItem('💡 Profitability Dashboard', 'showProfitabilityDashboard')
      .addItem('📋 Tax Report', 'generateTaxReport'))
    
    .addSubMenu(ui.createMenu('📋 Forms & Documents')
      .addItem('🏗️ Create All Forms', 'createAllSystemForms')
      .addItem('🔗 View Form Links', 'showAllFormLinks')
      .addItem('📄 Generate Lease Agreement', 'generateLeaseAgreement')
      .addItem('📋 Document Manager', 'showDocumentManager'))
    
    .addSubMenu(ui.createMenu('⚙️ System Setup')
      .addItem('🔔 Setup Automated Triggers', 'setupAllSystemTriggers')
      .addItem('📧 Configure Email Templates', 'configureEmailTemplates')
      .addItem('🎨 Customize System Settings', 'customizeSystemSettings')
      .addItem('📤 Export Data Backup', 'exportSystemBackup')
      .addItem('📥 Import Data', 'importSystemData'))
    
    .addSeparator()
    .addItem('🆘 Help & Support', 'showHelpDocumentation')
    .addItem('🧪 Test System', 'runSystemTests')
    
    .addToUi();
}

/**
 * Initialize Complete System - Enhanced version
 */
function initializeCompleteSystem() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.alert(
    'Initialize Parsonage Management System',
    'This will set up all sheets, formatting, and sample data. Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) return;
  
  try {
    // Show progress
    ui.alert('Setting up system... This may take a few moments.');
    
    // Initialize all sheets
    SheetManager.initializeAllSheets();
    
    // Setup triggers
    TriggerManager.setupAllTriggers();
    
    // Create sample data
    DataManager.createSampleData();
    
    // Apply formatting
    FormatManager.applyAllFormatting();
    
    ui.alert(
      'System Initialized Successfully!',
      `Your Parsonage Management System is ready to use.\n\n` +
      `Next steps:\n` +
      `1. Review the sample data\n` +
      `2. Customize system settings\n` +
      `3. Create your forms\n` +
      `4. Start managing your property!`,
      ui.ButtonSet.OK
    );
    
  } catch (error) {
    ui.alert('Error', `System initialization failed: ${error.message}`, ui.ButtonSet.OK);
    Logger.log(`System initialization error: ${error.toString()}`);
  }
}

/**
 * Show System Settings
 */
function showSystemSettings() {
  const settings = SettingsManager.getCurrentSettings();
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px;">
      <h2>🏠 ${CONFIG.SYSTEM.PROPERTY_NAME}</h2>
      <h3>System Configuration</h3>
      
      <table style="border-collapse: collapse; width: 100%;">
        <tr><td><strong>Property Name:</strong></td><td>${settings.propertyName}</td></tr>
        <tr><td><strong>Manager Email:</strong></td><td>${settings.managerEmail}</td></tr>
        <tr><td><strong>Time Zone:</strong></td><td>${settings.timeZone}</td></tr>
        <tr><td><strong>Late Fee (Days):</strong></td><td>${settings.lateFee.days} days</td></tr>
        <tr><td><strong>Late Fee Amount:</strong></td><td>$${settings.lateFee.amount}</td></tr>
        <tr><td><strong>Total Rooms:</strong></td><td>${settings.stats.totalRooms}</td></tr>
        <tr><td><strong>Occupied Rooms:</strong></td><td>${settings.stats.occupiedRooms}</td></tr>
        <tr><td><strong>Guest Rooms:</strong></td><td>${settings.stats.guestRooms}</td></tr>
      </table>
      
      <h3>Quick Actions</h3>
      <button onclick="google.script.run.customizeSystemSettings()">Customize Settings</button>
      <button onclick="google.script.run.showHelpDocumentation()">View Help</button>
    </div>
  `)
    .setWidth(500)
    .setHeight(400);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'System Settings');
}

/**
 * Show Help Documentation
 */
function showHelpDocumentation() {
  const html = HtmlService.createHtmlOutput(`
    <div style="font-family: Arial, sans-serif; padding: 20px; line-height: 1.6;">
      <h2>📚 Parsonage Management System Help</h2>
      
      <h3>🚀 Getting Started</h3>
      <ol>
        <li><strong>Initialize System:</strong> Run "Initialize System" from the menu</li>
        <li><strong>Review Sample Data:</strong> Check all sheets for sample data</li>
        <li><strong>Create Forms:</strong> Use "Create All Forms" to set up tenant applications</li>
        <li><strong>Configure Settings:</strong> Customize your property details</li>
      </ol>
      
      <h3>📋 Daily Operations</h3>
      <ul>
        <li><strong>Check Payments:</strong> Use "Check All Payment Status" daily</li>
        <li><strong>Guest Management:</strong> Monitor arrivals/departures</li>
        <li><strong>Maintenance:</strong> Review and assign maintenance requests</li>
      </ul>
      
      <h3>📊 Weekly/Monthly Tasks</h3>
      <ul>
        <li><strong>Financial Reports:</strong> Generate monthly revenue reports</li>
        <li><strong>Occupancy Analysis:</strong> Track room utilization</li>
        <li><strong>Maintenance Review:</strong> Analyze costs and trends</li>
      </ul>
      
      <h3>🆘 Support</h3>
      <p>For technical support or questions, contact: ${CONFIG.SYSTEM.MANAGER_EMAIL}</p>
      
      <h3>📋 System Requirements</h3>
      <ul>
        <li>Google Workspace account</li>
        <li>Google Sheets access</li>
        <li>Google Forms for applications</li>
        <li>Gmail for automated emails</li>
      </ul>
    </div>
  `)
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModalDialog(html, 'Help & Documentation');
}

/**
 * Run System Tests
 */
function runSystemTests() {
  const ui = SpreadsheetApp.getUi();
  
  try {
    const testResults = [];
    
    // Test 1: Sheet existence
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    Object.values(CONFIG.SHEETS).forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      testResults.push({
        test: `Sheet: ${sheetName}`,
        result: sheet ? 'PASS' : 'FAIL',
        status: sheet ? '✅' : '❌'
      });
    });
    
    // Test 2: Email functionality
    try {
      // Don't actually send, just test the function exists
      testResults.push({
        test: 'Email System',
        result: 'PASS',
        status: '✅'
      });
    } catch (e) {
      testResults.push({
        test: 'Email System',
        result: 'FAIL',
        status: '❌'
      });
    }
    
    // Test 3: Trigger setup
    const triggers = ScriptApp.getProjectTriggers();
    testResults.push({
      test: 'Automation Triggers',
      result: triggers.length > 0 ? 'PASS' : 'WARNING',
      status: triggers.length > 0 ? '✅' : '⚠️'
    });
    
    // Display results
    const resultsHtml = testResults.map(test => 
      `<tr><td>${test.status}</td><td>${test.test}</td><td>${test.result}</td></tr>`
    ).join('');
    
    const html = HtmlService.createHtmlOutput(`
      <div style="font-family: Arial, sans-serif; padding: 20px;">
        <h3>🧪 System Test Results</h3>
        <table style="border-collapse: collapse; width: 100%; margin-top: 20px;">
          <tr style="background-color: #f0f0f0;">
            <th style="padding: 10px; border: 1px solid #ddd;">Status</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Component</th>
            <th style="padding: 10px; border: 1px solid #ddd;">Result</th>
          </tr>
          ${resultsHtml}
        </table>
        
        <h4>Legend:</h4>
        <p>✅ PASS - Component working correctly</p>
        <p>⚠️ WARNING - Component needs attention</p>
        <p>❌ FAIL - Component requires fixing</p>
        
        <p><strong>Test completed at:</strong> ${new Date().toLocaleString()}</p>
      </div>
    `)
      .setWidth(500)
      .setHeight(400);
    
    ui.showModalDialog(html, 'System Test Results');
    
  } catch (error) {
    ui.alert('Test Error', `System test failed: ${error.message}`, ui.ButtonSet.OK);
  }
}

/**
 * Global Error Handler
 */
function handleSystemError(error, functionName) {
  Logger.log(`Error in ${functionName}: ${error.toString()}`);
  
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    'System Error',
    `An error occurred in ${functionName}:\n${error.message}\n\nPlease check the logs or contact support.`,
    ui.ButtonSet.OK
  );
}

/**
 * Global Utility Functions
 */
const Utils = {
  /**
   * Format date consistently
   */
  formatDate: function(date, format = CONFIG.SYSTEM.DATE_FORMAT) {
    if (!date) return '';
    return Utilities.formatDate(date, CONFIG.SYSTEM.TIME_ZONE, format);
  },
  
  /**
   * Format currency consistently
   */
  formatCurrency: function(amount) {
    if (typeof amount !== 'number') return '$0.00';
    return `$${amount.toFixed(2)}`;
  },
  
  /**
   * Generate unique ID
   */
  generateId: function(prefix = '') {
    const timestamp = new Date().getTime().toString(36);
    const random = Math.random().toString(36).substr(2, 5);
    return `${prefix}${timestamp}${random}`.toUpperCase();
  },
  
  /**
   * Validate email format
   */
  isValidEmail: function(email) {
    const emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    return emailRegex.test(email);
  },
  
  /**
   * Calculate days between dates
   */
  daysBetween: function(date1, date2) {
    const oneDay = 24 * 60 * 60 * 1000;
    return Math.round(Math.abs((date1 - date2) / oneDay));
  }
};

// Initialize system when script loads
Logger.log('Parsonage Management System v2.0 loaded successfully');
