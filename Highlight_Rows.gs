// -----------------------
// UI Functions
// -----------------------

function onOpen() {
  createCustomMenu();
}

function createCustomMenu() {
  var ui = SpreadsheetApp.getUi();
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
  var html = HtmlService.createHtmlOutputFromFile('ColorPicker')
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
  var email = Session.getActiveUser().getEmail();
  return email + "_" + key;
}

function getUserProperty(key) {
  var properties = PropertiesService.getScriptProperties();
  return properties.getProperty(getUserEmailPrefixedKey(key));
}

function setUserProperty(key, value) {
  var properties = PropertiesService.getScriptProperties();
  properties.setProperty(getUserEmailPrefixedKey(key), value);
}

function getHighlightProperties() {
  var rawProperties = getUserProperty('highlightProperties');
  return rawProperties ? JSON.parse(rawProperties) : {};
}

function setHighlightProperties(properties) {
  setUserProperty('highlightProperties', JSON.stringify(properties));
}

function setUserColorPreference(color) {
  var properties = getHighlightProperties();
  properties.highlightColor = color;
  setHighlightProperties(properties);
  showToastMessage('Row highlighting turned on. Selecting a cell will now highlight the entire row.');
}

function setHighlightingStatus(isEnabled) {
  var properties = getHighlightProperties();
  properties.rowHighlightEnabled = isEnabled ? 'true' : 'false';
  setHighlightProperties(properties);
}

// -----------------------
// Highlighting Functions
// -----------------------

function clearPreviouslyHighlightedRow() {
  var properties = getHighlightProperties();
  var previouslyHighlightedRow = properties.highlightedRow;
  var sheet = SpreadsheetApp.getActiveSheet();
  var maxColumns = sheet.getMaxColumns();

  if (previouslyHighlightedRow) {
    var originalColors = JSON.parse(properties.highlightedRowColors);
    if (originalColors && originalColors.length == maxColumns) {
      sheet.getRange(Number(previouslyHighlightedRow), 1, 1, maxColumns).setBackgrounds([originalColors]);
    }
  }
}

function onSelectionChange(e) {
  var properties = getHighlightProperties();
  var isHighlightingEnabled = properties.rowHighlightEnabled;

  if (isHighlightingEnabled !== 'true') {
    return;  // Exit the function if the row highlighting is turned off
  }

  var sheet = SpreadsheetApp.getActiveSheet();
  var currentRow = sheet.getActiveRange().getRow();
  var previouslyHighlightedRow = properties.highlightedRow;
  
  // If the currently selected row is the same as the previously highlighted row, exit the function
  if (previouslyHighlightedRow == currentRow) {
    return;
  }
  
  var maxColumns = sheet.getMaxColumns();
  var highlightColor = properties.highlightColor || '#DFF8FB'; // Default to Light Cyan if no color is set
  var highlightColors = Array(maxColumns).fill(highlightColor);

  // Save the original background colors of the current row to the properties
  var currentRowColors = sheet.getRange(currentRow, 1, 1, maxColumns).getBackgrounds();
  sheet.getRange(currentRow, 1, 1, maxColumns).setBackgrounds([highlightColors]);

  if (previouslyHighlightedRow) {
    restoreOriginalColorsForRow(previouslyHighlightedRow, maxColumns);
  }

  properties.highlightedRowColors = JSON.stringify(currentRowColors[0]);
  properties.highlightedRow = currentRow.toString();

  setHighlightProperties(properties);
}


function restoreOriginalColorsForRow(row, maxColumns) {
  var properties = getHighlightProperties();
  var originalColors = JSON.parse(properties.highlightedRowColors); // get from properties object
  var sheet = SpreadsheetApp.getActiveSheet();
  if (originalColors && originalColors.length == maxColumns) {
    sheet.getRange(Number(row), 1, 1, maxColumns).setBackgrounds([originalColors]);
  }
}

function unhighlightAllRows() {
  var properties = PropertiesService.getScriptProperties().getProperties();
  var sheet = SpreadsheetApp.getActiveSheet();
  var maxColumns = sheet.getMaxColumns();
  
  for (var propertyKey in properties) {
    // Check if the property is related to 'highlightedRow'
    if (propertyKey.endsWith("_highlightProperties")) {
      var userProperties = JSON.parse(properties[propertyKey]);
      
      // Check if there's a previously highlighted row for this user
      if (userProperties && userProperties.highlightedRow) {
        var previouslyHighlightedRow = userProperties.highlightedRow;

        // If the original colors for this row are stored, restore them
        if (userProperties.highlightedRowColors) {
          var originalColors = JSON.parse(userProperties.highlightedRowColors);
          if (originalColors && originalColors.length == maxColumns) {
            sheet.getRange(Number(previouslyHighlightedRow), 1, 1, maxColumns).setBackgrounds([originalColors]);
          }
        }

        // Clear the properties for this user
        userProperties.highlightedRow = null;
        userProperties.highlightedRowColors = null;
        PropertiesService.getScriptProperties().setProperty(propertyKey, JSON.stringify(userProperties));
      }
    }
  }

  showToastMessage('All highlighted rows have been cleared.');
}
