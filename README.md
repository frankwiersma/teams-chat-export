# Teams Chat Export

A Chrome/Edge extension to export Microsoft Teams chat conversations.

## Installation

1. Open Chrome or Edge
2. Go to `chrome://extensions` (Chrome) or `edge://extensions` (Edge)
3. Enable **Developer mode** (toggle in top right)
4. Click **Load unpacked**
5. Select the extension folder

## Usage

1. Open a Teams chat at https://teams.microsoft.com
2. Click the extension icon in the toolbar
3. Click **Scroll to Top** to load the full chat history
4. Click one of the export buttons: **HTML**, **Text**, **JSON**, or **CSV**

## Features

- Exports full chat history with sender names and timestamps
- Uses the chat name as the export filename
- Multiple export formats:
  - **HTML**: Styled, readable in any browser
  - **Text**: Plain text, easy to read
  - **JSON**: Structured data for processing
  - **CSV**: Excel-compatible spreadsheet

## Supported

- Microsoft Teams (new version): teams.microsoft.com/v2/
- Works in Chrome and Edge

## Files

```
teams-chat-export/
├── manifest.json    # Extension manifest
├── popup.html       # Extension popup UI
├── popup.js         # Popup logic
├── content.js       # Content script (runs on Teams)
└── icons/
    ├── icon16.png
    ├── icon48.png
    └── icon128.png
```
