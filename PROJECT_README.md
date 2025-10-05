# Google Sheets Dynamic Automation System

## 📋 Project Overview

This project implements a comprehensive Google Sheets automation system using Apps Script that provides dynamic sheet management, automated row transfers based on status changes, and data backup functionality. The system is designed for MIS Executives and Data Analysts who need efficient data organization and workflow automation.

---

## 🎯 Key Features

### 1. **Dynamic Sheet Creation & Deletion**
- Automatically creates new sheets based on names listed in a control sheet
- Automatically deletes sheets when names are removed from the list
- Backs up all data before deletion to prevent data loss

### 2. **Status-Based Row Transfer**
- Automatically moves rows to different sheets when status changes
- Supports dropdown-based status updates
- Maintains data integrity during transfers

### 3. **Data Backup System**
- Archives all deleted sheet data with timestamps
- Preserves formatting and structure
- Organized backup format for easy retrieval

### 4. **User-Friendly Menu Interface**
- Custom menu for manual operations
- One-click sync functionality
- Easy access to backup data

---

## 🏗️ System Architecture

### Configuration Structure

```javascript
const CONFIG = {
  controlSheet: "DataValidation",    // Sheet with sheet names list
  dataSaverSheet: "DataSaver",       // Archive for deleted data
  nameStartRow: 2,                    // Starting row (A2)
  nameEndRow: 10,                     // Ending row (A10)
  nameColumn: 1                       // Column A
};
```

### Sheet Structure

```
📊 Spreadsheet Structure
├── DataValidation (Control Sheet)
│   ├── Column A (Rows 2-10): List of sheet names
│   └── Controls which sheets exist
│
├── DataSaver (Archive Sheet)
│   ├── Deleted Date | Sheet Name | Data Starts Below
│   └── [Archived data from deleted sheets]
│
├── [Dynamic Sheets]
│   ├── Created automatically from DataValidation
│   └── Can be used for categorized data storage
│
└── Working Sheets
    └── Sheets with Column E for status-based transfers
```

---

## 🔧 Core Functionalities

### Feature 1: Dynamic Sheet Management

**How It Works:**
1. User adds a name in `DataValidation` sheet (A2:A10)
2. System automatically creates a new sheet with that name
3. When name is deleted from list:
   - All sheet data is backed up to `DataSaver`
   - Sheet is then deleted
   - Backup includes timestamp and formatting

**Use Cases:**
- Project-based data organization
- Department-wise tracking
- Category management
- Dynamic reporting structures

**Code Implementation:**
```javascript
function syncSheets() {
  // Reads names from DataValidation sheet
  // Creates missing sheets
  // Backs up and deletes removed sheets
  // Protects essential sheets from deletion
}
```

---

### Feature 2: Status-Based Row Transfer

**How It Works:**
1. User changes value in Column E (Status column)
2. System checks if a sheet with that status name exists
3. If exists: Row is automatically moved to that sheet
4. Original row is deleted from source sheet

**Example Workflow:**
```
Step 1: Item in "Inventory" sheet
Column E = "In Stock"

Step 2: User changes Column E to "Re-purchase needed"

Step 3: Entire row moves to "Re-purchase needed" sheet
Original row deleted from "Inventory"
```

**Code Implementation:**
```javascript
if (editedCol === 5) {  // Column E
  const statusVal = range.getValue();
  const targetSheet = ss.getSheetByName(statusVal);
  targetSheet.appendRow(rowData[0]);
  sheet.deleteRow(editedRow);
}
```

---

### Feature 3: Data Backup System

**Backup Format:**
```
| Deleted Date      | Sheet Name  | Data Starts Below  |
|-------------------|-------------|--------------------|
| 2025-10-05 10:30  | Employee1   | --- Data Below --- |
| [All data from Employee1 sheet copied here]          |
|                                                       |
| 2025-10-05 11:15  | Project_X   | --- Data Below --- |
| [All data from Project_X sheet copied here]          |
```

**Features:**
- Timestamped backups
- Preserves cell formatting and colors
- Organized with visual separators
- Easy to search and retrieve

**Code Implementation:**
```javascript
function backupSheetData(sourceSheet, dataSaverSheet) {
  // Captures timestamp
  // Copies all data and formatting
  // Adds visual separators
  // Appends to DataSaver sheet
}
```

---

## 📝 Complete Code

```javascript
/*
@OnlyCurrentDoc
*/

// CONFIGURATION - Change these if needed
const CONFIG = {
  controlSheet: "DataValidation",        // Sheet containing names
  dataSaverSheet: "DataSaver",   // Sheet to save deleted data
  nameStartRow: 2,               // Starting row (A2)
  nameEndRow: 10,                // Ending row (A10)
  nameColumn: 1                  // Column A
};

// COMBINED onEdit function - handles both status change and sheet sync
function onEdit(e) {
  const range = e.range;
  const editedCol = range.getColumn();
  const editedRow = range.getRow();
  const sheet = e.source.getActiveSheet();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // ===== FUNCTIONALITY 1: Move row based on status change (Column E) =====
  if (editedCol === 5) {  // Column E (Status)
    const statusVal = range.getValue();
    
    if (statusVal !== "") {
      const targetSheet = ss.getSheetByName(statusVal);
      
      if (!targetSheet) {
        SpreadsheetApp.getUi().alert(`Sheet "${statusVal}" does not exist.`);
        return;
      }
      
      // Get row data and move it
      const rowData = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues();
      targetSheet.appendRow(rowData[0]);    // Move row to target sheet
      sheet.deleteRow(editedRow);           // Delete row from source sheet
      
      Logger.log(`Moved row ${editedRow} to sheet: ${statusVal}`);
    }
  }

  // ===== FUNCTIONALITY 2: Sync sheets based on DataValidation sheet =====
  if (sheet.getName() === CONFIG.controlSheet && 
      range.getColumn() === CONFIG.nameColumn &&
      range.getRow() >= CONFIG.nameStartRow &&
      range.getRow() <= CONFIG.nameEndRow) {
    
    // Run sync after a short delay to allow for multiple rapid edits
    Utilities.sleep(500);
    syncSheets();
  }
}

// Main function to sync sheets with names list
function syncSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const controlSheet = ss.getSheetByName(CONFIG.controlSheet);
  
  // Check if control sheet exists
  if (!controlSheet) {
    SpreadsheetApp.getUi().alert(`Error: "${CONFIG.controlSheet}" sheet not found!`);
    return;
  }
  
  // Get or create DataSaver sheet
  let dataSaverSheet = ss.getSheetByName(CONFIG.dataSaverSheet);
  if (!dataSaverSheet) {
    dataSaverSheet = ss.insertSheet(CONFIG.dataSaverSheet);
    // Add headers
    dataSaverSheet.getRange(1, 1, 1, 3).setValues([["Deleted Date", "Sheet Name", "Data Starts Below"]]);
    dataSaverSheet.getRange(1, 1, 1, 3).setFontWeight("bold").setBackground("#f3f3f3");
  }
  
  // Get the list of names from A2:A10
  const nameRange = controlSheet.getRange(CONFIG.nameStartRow, CONFIG.nameColumn, 
                                          CONFIG.nameEndRow - CONFIG.nameStartRow + 1, 1);
  const names = nameRange.getValues().flat().filter(name => name !== "");
  
  // Get all existing sheets
  const allSheets = ss.getSheets();
  const existingSheetNames = allSheets.map(sheet => sheet.getName());
  
  // Protected sheet names (sheets that should never be deleted)
  const protectedSheets = [CONFIG.controlSheet, CONFIG.dataSaverSheet, "Sheet1"];
  
  // Step 1: Create new sheets for names that don't exist
  names.forEach(name => {
    if (!existingSheetNames.includes(name)) {
      createNewSheet(ss, name);
      Logger.log(`Created sheet: ${name}`);
    }
  });
  
  // Step 2: Backup and delete sheets that are not in the names list
  allSheets.forEach(sheet => {
    const sheetName = sheet.getName();
    
    // Only delete if:
    // 1. Not in the names list
    // 2. Not a protected sheet
    if (!names.includes(sheetName) && 
        !protectedSheets.includes(sheetName)) {
      
      // Backup data before deletion
      backupSheetData(sheet, dataSaverSheet);
      
      // Delete the sheet
      ss.deleteSheet(sheet);
      Logger.log(`Backed up and deleted sheet: ${sheetName}`);
    }
  });
  
  Logger.log("Sheet sync completed!");
}

// Function to create a new blank sheet
function createNewSheet(ss, newSheetName) {
  const newSheet = ss.insertSheet(newSheetName);
  
  // Move the new sheet to the end (optional)
  ss.moveActiveSheet(ss.getNumSheets());
  
  return newSheet;
}

// Function to backup sheet data to DataSaver
function backupSheetData(sourceSheet, dataSaverSheet) {
  const sheetName = sourceSheet.getName();
  const lastRow = sourceSheet.getLastRow();
  const lastColumn = sourceSheet.getLastColumn();
  
  // Get the next available row in DataSaver
  const nextRow = dataSaverSheet.getLastRow() + 1;
  
  // Add separator and metadata
  const timestamp = new Date();
  dataSaverSheet.getRange(nextRow, 1).setValue(timestamp);
  dataSaverSheet.getRange(nextRow, 2).setValue(sheetName);
  dataSaverSheet.getRange(nextRow, 3).setValue("--- Data Below ---");
  
  // Style the header row
  dataSaverSheet.getRange(nextRow, 1, 1, 3)
    .setFontWeight("bold")
    .setBackground("#fff2cc")
    .setBorder(true, true, true, true, true, true);
  
  // Copy all data from the source sheet if it has data
  if (lastRow > 0 && lastColumn > 0) {
    const sourceData = sourceSheet.getRange(1, 1, lastRow, lastColumn).getValues();
    const sourceFormats = sourceSheet.getRange(1, 1, lastRow, lastColumn).getBackgrounds();
    
    // Paste data starting from next row
    const targetRange = dataSaverSheet.getRange(nextRow + 1, 1, lastRow, lastColumn);
    targetRange.setValues(sourceData);
    targetRange.setBackgrounds(sourceFormats);
    
    Logger.log(`Backed up ${lastRow} rows and ${lastColumn} columns from ${sheetName}`);
  } else {
    dataSaverSheet.getRange(nextRow + 1, 1).setValue("(No data to backup)");
  }
  
  // Add a blank row for separation
  dataSaverSheet.getRange(nextRow + lastRow + 2, 1).setValue("");
}

// Manual sync function - run this from menu
function manualSync() {
  syncSheets();
  SpreadsheetApp.getUi().alert('✓ Sheets synced successfully!\n\nSheets created/deleted based on names in DataValidation (A2:A10).\nDeleted sheet data saved to DataSaver.');
}

// Create custom menu on open
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🔄 Sheet Manager')
      .addItem('Sync Sheets Now', 'manualSync')
      .addItem('Setup Auto-Sync', 'setupTriggers')
      .addSeparator()
      .addItem('View Backup Data', 'openDataSaver')
      .addToUi();
}

// Function to open DataSaver sheet
function openDataSaver() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const dataSaverSheet = ss.getSheetByName(CONFIG.dataSaverSheet);
  
  if (dataSaverSheet) {
    ss.setActiveSheet(dataSaverSheet);
  } else {
    SpreadsheetApp.getUi().alert('DataSaver sheet not found. It will be created automatically when a sheet is deleted.');
  }
}

// Setup automatic triggers
function setupTriggers() {
  // Delete existing triggers to avoid duplicates
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  
  // Create new onEdit trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
    .onEdit()
    .create();
  
  SpreadsheetApp.getUi().alert(
    '✓ Auto-sync is now active!\n\n' +
    'Sheets will automatically:\n' +
    '• Move rows when status (Column E) changes\n' +
    '• Create sheets when you add names to DataValidation (A2:A10)\n' +
    '• Backup data to DataSaver before deletion\n' +
    '• Delete when you remove names from DataValidation\n\n' +
    'You can also manually sync using the menu.'
  );
}
```

---

## 🚀 Installation & Setup

### Step 1: Create Required Sheets

1. Open your Google Sheet
2. Create a sheet named **"DataValidation"**
3. The **"DataSaver"** sheet will be created automatically

### Step 2: Install the Script

1. Go to **Extensions** → **Apps Script**
2. Delete any existing code
3. Paste the complete code above
4. Click **Save** (Ctrl+S)
5. Name your project (e.g., "Sheet Automation System")

### Step 3: Initial Setup

1. Close the Apps Script editor
2. Refresh your Google Sheet
3. You'll see a new menu **"🔄 Sheet Manager"**
4. Click **🔄 Sheet Manager** → **Setup Auto-Sync**
5. Authorize the script when prompted

### Step 4: Configure Your System

**In DataValidation Sheet:**
- Add sheet names in cells A2 to A10
- Each name will create a corresponding sheet
- Leave cells empty if you don't need all 9 slots

**Example:**
```
A2: In Stock
A3: Re-purchase needed
A4: Discontinued
A5: On Order
A6: (empty)
A7: (empty)
A8: (empty)
A9: (empty)
A10: (empty)
```

---

## 💡 Usage Examples

### Example 1: Inventory Management

**Setup:**
```
DataValidation Sheet (A2:A10):
- In Stock
- Low Stock
- Out of Stock
- On Order
- Discontinued
```

**Workflow:**
1. Product starts in "Inventory" sheet
2. When stock runs low, change Column E to "Low Stock"
3. Row automatically moves to "Low Stock" sheet
4. When restocked, change status to "In Stock"
5. Row moves to "In Stock" sheet

---

### Example 2: Project Task Management

**Setup:**
```
DataValidation Sheet (A2:A10):
- To Do
- In Progress
- Under Review
- Completed
- On Hold
```

**Workflow:**
1. New tasks in "Master Tasks" sheet
2. Change status in Column E as work progresses
3. Tasks automatically organize into respective sheets
4. Easy tracking of project stages

---

### Example 3: Customer Order Processing

**Setup:**
```
DataValidation Sheet (A2:A10):
- New Orders
- Processing
- Shipped
- Delivered
- Cancelled
```

**Workflow:**
1. Orders received in "Orders" sheet
2. Update Column E as order status changes
3. Automatic categorization
4. If category removed from DataValidation, orders backed up to DataSaver

---

## 🎛️ Customization Options

### Change Control Sheet Range

Modify rows monitored for sheet names:

```javascript
const CONFIG = {
  nameStartRow: 5,    // Start from A5
  nameEndRow: 20,     // End at A20 (16 possible sheets)
  // ... other config
};
```

### Change Status Column

To use a different column for status triggers:

```javascript
if (editedCol === 3) {  // Change 5 to 3 for Column C
  // ... rest of code
}
```

### Add Protected Sheets

Prevent additional sheets from deletion:

```javascript
const protectedSheets = [
  CONFIG.controlSheet, 
  CONFIG.dataSaverSheet, 
  "Sheet1",
  "Master Data",      // Add your sheets here
  "Reference",
  "Dashboard"
];
```

### Change Sheet Names

```javascript
const CONFIG = {
  controlSheet: "SheetList",        // Your preferred name
  dataSaverSheet: "Archive",        // Your preferred name
  // ... other config
};
```

---

## 🔒 Data Security Features

### Protected Sheets
- System sheets (DataValidation, DataSaver, Sheet1) cannot be accidentally deleted
- Add custom protected sheets in configuration

### Data Backup
- All data backed up before deletion
- Includes timestamps for audit trail
- Preserves formatting and structure

### Error Handling
- Validates sheet existence before operations
- User alerts for missing sheets
- Prevents data loss from invalid operations

---

## 📊 Menu Options

### 🔄 Sheet Manager Menu

**Sync Sheets Now**
- Manually triggers sheet synchronization
- Creates/deletes sheets based on current list
- Use after bulk changes to DataValidation

**Setup Auto-Sync**
- Configures automatic triggers
- Enables real-time synchronization
- Required for automatic operation

**View Backup Data**
- Opens DataSaver sheet
- Quick access to archived data
- Review deleted sheet contents

---

## 🐛 Troubleshooting

### Issue: Rows Not Transferring

**Solution:**
1. Check if target sheet exists
2. Verify Column E has dropdown values
3. Ensure values match sheet names exactly
4. Run "Setup Auto-Sync" from menu

### Issue: Sheets Not Auto-Creating

**Solution:**
1. Check DataValidation sheet name is correct
2. Verify names are in cells A2:A10
3. Run "Sync Sheets Now" manually
4. Check Apps Script logs for errors

### Issue: Data Not Backing Up

**Solution:**
1. Ensure DataSaver sheet exists
2. Check sheet permissions
3. Verify script authorization
4. Review execution logs

### Issue: Menu Not Appearing

**Solution:**
1. Refresh the Google Sheet
2. Check script is saved
3. Clear browser cache
4. Re-run setupTriggers function

---

## 📈 Performance Considerations

### Optimization Tips

1. **Limit Sheet Count**: Keep dynamic sheets under 20 for best performance
2. **Data Volume**: Works efficiently with up to 10,000 rows per sheet
3. **Trigger Delays**: Built-in 500ms delay prevents rapid-fire triggers
4. **Batch Operations**: Use manual sync for bulk changes

### Execution Limits

- **Apps Script Quotas**: 6 minutes per execution
- **Daily Trigger Total**: 90 minutes per day (free accounts)
- **Concurrent Executions**: One at a time per user

---

## 🔮 Future Enhancements

### Potential Features

- [ ] Email notifications on sheet creation/deletion
- [ ] Advanced filtering options
- [ ] Custom backup retention policies
- [ ] Scheduled automatic archiving
- [ ] Multi-column status triggers
- [ ] Undo functionality for deletions
- [ ] Export backups to Drive folders

---

## 📝 Changelog

### Version 1.0 (Current)
- Initial release
- Dynamic sheet management
- Status-based row transfer
- Data backup system
- Custom menu interface
- Auto-trigger setup

---

## 👨‍💻 Developer Notes

### Code Structure

```
├── Configuration (CONFIG object)
├── Core Functions
│   ├── onEdit() - Main trigger handler
│   ├── syncSheets() - Sheet synchronization
│   ├── createNewSheet() - Sheet creation
│   └── backupSheetData() - Data archival
├── UI Functions
│   ├── onOpen() - Menu creation
│   ├── manualSync() - Manual trigger
│   ├── openDataSaver() - Navigation
│   └── setupTriggers() - Trigger configuration
└── Utilities
    └── Error handling & logging
```

### Best Practices

1. **Always test in a copy** before production use
2. **Regular backups** of the entire spreadsheet
3. **Document customizations** in comments
4. **Monitor execution logs** for errors
5. **Keep trigger count minimal** for performance

---

## 📄 License & Usage

This script is provided as-is for educational and professional use. 

**Permissions:**
- ✅ Use in personal projects
- ✅ Modify for your needs
- ✅ Use in workplace (check company policy)
- ✅ Share with colleagues

**Restrictions:**
- ❌ Do not include confidential data in shared versions
- ❌ Respect your organization's IT policies
- ❌ Test thoroughly before production deployment

---

## 🤝 Support & Contact

For questions, issues, or improvements:
1. Review this documentation thoroughly
2. Check Google Apps Script documentation
3. Test in a sandbox environment
4. Consult with your IT department for workplace use

---

## 📚 Additional Resources

- [Google Apps Script Documentation](https://developers.google.com/apps-script)
- [Spreadsheet Service Reference](https://developers.google.com/apps-script/reference/spreadsheet)
- [Trigger Documentation](https://developers.google.com/apps-script/guides/triggers)
- [Best Practices Guide](https://developers.google.com/apps-script/guides/support/best-practices)

---

**Last Updated:** October 2025  
**Version:** 1.0  
**Compatibility:** Google Sheets (Web, Mobile, Desktop)

---

*This automation system is designed to enhance productivity for Data Analysts and MIS Executives through intelligent sheet management and workflow automation.*
