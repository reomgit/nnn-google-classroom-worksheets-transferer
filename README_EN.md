# Classroom Discovery Tool

> A Google Apps Script that finds your "Classroom" folder automatically, lists your classes, and lets you manage files — all without leaving Google Sheets.

---

## Table of Contents

- [Features](#features)
- [Setup](#setup)
- [Usage](#usage)
- [Configuration](#configuration)
- [How It Works](#how-it-works)
- [Troubleshooting](#troubleshooting)

---

## Features

- Automatically discovers your Google Classroom folder in Drive
- Lists all class subfolders with checkboxes for selection
- Fetches and sorts files from any selected class folder
- Resolves Google Drive shortcuts to their real target files
- Moves files to destination folders via pasted URLs
- Syncs file names to match shortcut display names
- Ownership safety check — only moves files you own

---

## Setup

1. Create a new Google Sheet (e.g., "Classroom Manager").
2. Go to **Extensions > Apps Script**.
3. **Copy the code from [`Code_EN.gs`](./Code_EN.gs)**.
4. Delete any existing code in the editor and paste the copied code.
5. Save the project.
6. Refresh the Google Sheet.

---

## Usage

### 1. Authorize

Run any command from the menu for the first time. You'll see an "Authorization Required" popup — this is safe and only requests access to your own Drive folders.

### 2. Find My Classes

1. Click **Classroom Manager > 1. Find My Classes**.
2. The sheet populates with all subfolders inside your "Classroom" folder.
3. Each row has a checkbox, the class name, folder ID, and folder URL.

### 3. Select & Fetch Files

1. Check the box (Column A) next to the class you want to organize.
2. Click **2. Fetch Files from Selected Class**.
3. The sheet reloads with all files in that class, sorted by:
   - Files starting with `YYYY_` (natural numeric sort)
   - Remaining files (alphabetical, Japanese locale)

### 4. Assign & Move

1. Paste destination folder URLs into **Column C** for each file you want to move.
2. Click **3. Move Files**.
3. The script moves each file and updates the Status column.

---

## Configuration

Edit the variable at the top of the script to match your folder name:

```javascript
var ROOT_FOLDER_NAME = "Classroom";
```

Change `"Classroom"` to whatever your Google Classroom folder is actually named (e.g., `"Google Classroom"`).

---

## How It Works

| Function | Description |
|---|---|
| `onOpen()` | Creates the custom menu in Google Sheets |
| `listClassFolders()` | Searches Drive for the root folder and lists all subfolders |
| `fetchFilesFromSelection()` | Reads files from the selected class folder, resolves shortcuts, and sorts results |
| `processMoveQueue()` | Moves files to target folders and optionally renames them to match display names |
| `extractIdFromUrl()` | Extracts a Google Drive folder/file ID from a URL |

### Shortcut Handling

When the script encounters a Google Drive shortcut:

1. **Name** — keeps the shortcut's display name
2. **ID** — uses the real target file's ID (so moves affect the actual file)
3. **URL** — uses the target file's URL if accessible, otherwise falls back to the shortcut URL
4. Broken shortcuts (where `getTargetId()` fails) are skipped entirely

### Safety

- Only files **you own** are moved. Non-owned files get status `Skipped (Not Owner)`.
- Files are renamed only if the sheet name differs from the real file name.

---

## Troubleshooting

| Problem                                     | Solution                                                                   |
| ------------------------------------------- | -------------------------------------------------------------------------- |
| "Could not find a folder named 'Classroom'" | Change `ROOT_FOLDER_NAME` to match your actual folder name                 |
| Folder found but empty                      | Verify the folder has subfolders (not just files at the root level)        |
| Files marked `[Restricted Target]`          | The shortcut points to a file you don't have permission to access directly |
| `Skipped (Not Owner)`                       | You can only move files you own                                            |
