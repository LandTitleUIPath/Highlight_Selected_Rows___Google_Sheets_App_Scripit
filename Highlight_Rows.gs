// -----------------------
// UI Functions
// -----------------------

function onOpen() {
  createCustomMenu();
}

function createCustomMenu() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Auto Highlight Selected Row')
      .addItem('Turn On Row Highlight', 'turnOnRowHighlight')
      .addItem('Turn Off Row Highlight', 'turnOffRowHighlight')
      .addItem('Clear All User Highlights', 'unhighlightAllRows')
      .addToUi();
}

function showToastMessage(message) {
  SpreadsheetApp.getActiveSpreadsheet().toast(message, 'Status', 4);
}

function turnOnRowHighlight() {
  setHighlightingStatus(true);

  // Show color picker dialog
  const html = HtmlService.createHtmlOutputFromFile('ColorPickerModal')
      .setWidth(800)
      .setHeight(300);
  SpreadsheetApp.getUi().showModalDialog(html, 'Select a Highlight Color');
}

function turnOffRowHighlight() {
  setHighlightingStatus(false);
  clearPreviouslyHighlightedRow();
  unhighlightAllRows();
  showToastMessage('Row highlighting turned off. Your last highlighted row has been cleared.');
}

// -----------------------
// Property Management
// -----------------------

function getUserEmailPrefixedKey(key) {
  const email = Session.getActiveUser().getEmail();
  return email + "_" + key;
}

function getUserProperty(key) {
  const properties = PropertiesService.getScriptProperties();
  return properties.getProperty(getUserEmailPrefixedKey(key));
}

function setUserProperty(key, value) {
  const properties = PropertiesService.getScriptProperties();
  properties.setProperty(getUserEmailPrefixedKey(key), value);
}

function getHighlightProperties() {
  const rawProperties = getUserProperty('highlightProperties');
  return rawProperties ? JSON.parse(rawProperties) : {};
}

function setHighlightProperties(properties) {
  setUserProperty('highlightProperties', JSON.stringify(properties));
}

function setUserColorPreference(color) {
  let properties = getHighlightProperties();
  properties.highlightColor = color;
  setHighlightProperties(properties);
  showToastMessage('Row highlighting turned on. Selecting a cell will now highlight the entire row.');
}

function setHighlightingStatus(isEnabled) {
  let properties = getHighlightProperties();
  properties.rowHighlightEnabled = isEnabled ? 'true' : 'false';
  setHighlightProperties(properties);
}

// -----------------------
// Highlighting Functions
// -----------------------
function clearPreviouslyHighlightedRow(previouslyHighlightedRow, sheetName, originalColors) {
    const sheet = sheetName ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) : SpreadsheetApp.getActiveSheet();
    const maxColumns = sheet.getMaxColumns();

    if (previouslyHighlightedRow) {
        if (originalColors && originalColors.length == maxColumns) {
            sheet.getRange(Number(previouslyHighlightedRow), 1, 1, maxColumns).setBackgrounds([originalColors]);
        }
    }
}

function onSelectionChange(e) {
  const properties = getHighlightProperties();
  const isHighlightingEnabled = properties.rowHighlightEnabled;

  if (isHighlightingEnabled !== 'true') {
      return;  
  }

  const sheet = e.source.getActiveSheet(); // Use event object
  const currentRow = e.range.getRow();     // Use event object
  const previouslyHighlightedRow = properties.highlightedRow;

  if (previouslyHighlightedRow && previouslyHighlightedRow != currentRow) {
      clearPreviouslyHighlightedRow(previouslyHighlightedRow, properties.highlightedSheetName, JSON.parse(properties.highlightedRowColors));
  }

  if (previouslyHighlightedRow == currentRow) {
      return;
  }

  const maxColumns = sheet.getMaxColumns();
  const highlightColor = properties.highlightColor || '#DFF8FB';
  const highlightColors = Array(maxColumns).fill(highlightColor);

  const currentRowColors = e.range.offset(0, 0, 1, maxColumns).getBackgrounds(); // Use offset with event range
  sheet.getRange(currentRow, 1, 1, maxColumns).setBackgrounds([highlightColors]);

  properties.highlightedRowColors = JSON.stringify(currentRowColors[0]);
  properties.highlightedRow = currentRow.toString();
  properties.highlightedSheetName = sheet.getName();

  setHighlightProperties(properties);
}

function restoreOriginalColorsForRow(row, maxColumns) {
  const properties = getHighlightProperties();
  const originalColors = JSON.parse(properties.highlightedRowColors);
  const sheetName = properties.highlightedSheetName; 
  const sheet = sheetName ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) : SpreadsheetApp.getActiveSheet();

  if (originalColors && originalColors.length == maxColumns) {
    sheet.getRange(Number(row), 1, 1, maxColumns).setBackgrounds([originalColors]);
  }
}

function unhighlightAllRows() {
  const properties = PropertiesService.getScriptProperties().getProperties();
  const activeSheet = SpreadsheetApp.getActiveSheet();
  const maxColumns = activeSheet.getMaxColumns();
  
  for (let propertyKey in properties) {
    if (propertyKey.endsWith("_highlightProperties")) {
      const userProperties = JSON.parse(properties[propertyKey]);
      
      if (userProperties && userProperties.highlightedRow) {
        const sheetName = userProperties.highlightedSheetName; 
        const sheet = sheetName ? SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName) : activeSheet;
        const previouslyHighlightedRow = userProperties.highlightedRow;

        if (userProperties.highlightedRowColors) {
          const originalColors = JSON.parse(userProperties.highlightedRowColors);
          if (originalColors && originalColors.length == maxColumns) {
            sheet.getRange(Number(previouslyHighlightedRow), 1, 1, maxColumns).setBackgrounds([originalColors]);
          }
        }

        userProperties.highlightedRow = null;
        userProperties.highlightedRowColors = null;
        userProperties.highlightedSheetName = null;  // Clear the sheet name
        PropertiesService.getScriptProperties().setProperty(propertyKey, JSON.stringify(userProperties));
      }
    }
  }

  showToastMessage('All highlighted rows have been cleared.');
}
