# Auto Highlight Rows for Google Sheets

This script extension enhances Google Sheets with a capability that allows users to automatically highlight an entire row based on their current cell selection. Additionally, the users can choose from a palette of colors for the highlighting, ensuring the perfect match with their existing content.

## Table of Contents

- [Features](#features)
- [Installation](#installation)
- [Usage](#usage)
- [Files and Architecture](#files-and-architecture)

## Features

- **Custom Menu Integration**: Directly built into Google Sheets' menu, allowing for quick access to row highlighting features without needing to jump to another interface.
- **Toggle Ability**: Enables users to activate or deactivate the row highlighting feature with just a click.
- **Color Picker Modal**: A visually appealing interface that presents a selection of colors for users to choose for row highlighting.
- **State Persistence**: The script remembers the user's color choice and the last highlighted row, ensuring consistency across sessions.

## Installation

1. Open the desired Google Sheet where you wish to implement this feature.
2. Navigate to `Extensions > Apps Script`.
3. Erase any existing code and paste the content from `Highlight_Rows.gs` into the script editor.
4. In the script editor, create a new HTML file: `File > New > HTML File`. Name this file `ColorPickerModal`.
5. Paste the content from `ColorPickerModal.html` into this new file.
6. Save all changes and close the script editor.
7. Reload your Google Sheet. After a few seconds, you should notice a custom menu titled "Auto Highlight Selected Row".

## Usage

1. **Activation**: Click on the `Auto Highlight Selected Row` option in your Google Sheets menu. From the dropdown, select `Turn On Row Highlight`. A modal will pop up.
2. **Choosing a Color**: In the modal, click on your desired highlight color. This color will be used to highlight rows when you select any cell within them.
3. **Highlighting**: After choosing a color, simply click on any cell in your Google Sheet. The entire row containing the selected cell will be highlighted with the chosen color.
4. **Deactivation**: To turn off the auto-highlighting feature, go to the `Auto Highlight Selected Row` menu and select `Turn Off Row Highlight`.
5. **Clear Highlights**: To remove all row highlights, select `Clear All User Highlights` from the custom menu.

## Files and Architecture

- **Highlight_Rows.gs**: This file contains the core logic of the script, handling Google Sheet interactions, properties management, UI functions, and row highlighting functionalities.
- **ColorPickerModal.html**: An HTML file rendering a color picker UI, allowing users to select their desired row highlight color.

