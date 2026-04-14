# gdrive_export

Export a Google Sheet, Doc, or Slide to a file and upload it to a Drive folder.

| Source type | Output format |
|---|---|
| Google Sheets | `.xlsx` |
| Google Docs | `.pdf` |
| Google Slides | `.pptx` |

---

## Setup

### 1. Install dependencies

```bash
npm install
```

### 2. Create OAuth credentials

1. Go to [Google Cloud Console](https://console.cloud.google.com/)
2. Create a project (or select an existing one)
3. Enable the **Google Drive API**
4. Go to **APIs & Services → Credentials → Create Credentials → OAuth 2.0 Client ID**
5. Application type: **Desktop app**
6. Download the JSON file

### 3. Save the credentials

```bash
mkdir -p ~/.config/gdrive_export && mv credentials.json ~/.config/gdrive_export/credentials.json
```

### 4. Authorize

Run the script once — a browser window will open for Google sign-in. The token is cached at `~/.config/gdrive_export/token.json` for all subsequent runs.

---

## Usage

```
node gdrive_export.js <source> <folder> [--append-timestamp] [--filename NAME] [--dry-run]
```

`<source>` and `<folder>` accept either a full Google Drive URL or a bare file/folder ID.

### Examples

```bash
# Export a Sheet to a folder (using URLs)
node gdrive_export.js \
  "https://docs.google.com/spreadsheets/d/<FILE_ID>/edit" \
  "https://drive.google.com/drive/folders/<FOLDER_ID>"

# Use raw IDs
node gdrive_export.js <FILE_ID> <FOLDER_ID>

# Append a UTC timestamp to the filename
node gdrive_export.js <FILE_ID> <FOLDER_ID> --append-timestamp

# Override the output filename
node gdrive_export.js <FILE_ID> <FOLDER_ID> --filename "Q1 Report"

# Override name and append timestamp
node gdrive_export.js <FILE_ID> <FOLDER_ID> --filename "Q1 Report" --append-timestamp

# Preview without downloading or uploading
node gdrive_export.js <FILE_ID> <FOLDER_ID> --dry-run
```

### Options

| Flag | Description |
|---|---|
| `--append-timestamp` | Append a UTC timestamp suffix (`YYYYMMDD_HHMMSSz`) to the filename |
| `--filename NAME` | Override the base filename (without extension) instead of using the source file's Drive name |
| `--dry-run` | Print what would happen without touching Drive |

---

## File locations

| Path | Purpose |
|---|---|
| `~/.config/gdrive_export/credentials.json` | OAuth client credentials (you provide this) |
| `~/.config/gdrive_export/token.json` | Cached access token (auto-generated on first run) |

---

## Notes

- Works with files in **Shared Drives** as well as personal Drive.
- To re-authorize (e.g. after revoking access), delete `~/.config/gdrive_export/token.json` and re-run.
