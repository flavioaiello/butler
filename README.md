# Butler - Outlook Email Archiver

A Microsoft Edge/Chrome extension that automatically archives obsolete emails from conversation threads in Outlook Web App (OWA), keeping only the latest message (head) of each conversation. Also detects and moves duplicate emails.

## Features

- **Dry Run** - Preview which emails would be archived and duplicates detected
- **Archive All** - Move obsolete conversation emails to your Online Archive folder
- **Duplicate Detection** - Find and move duplicate emails (same Message-ID) to a "Duplicates" folder
- **Automatic Token Capture** - Seamlessly captures authentication from OWA

## How It Works

Butler keeps your inbox clean by:
1. Archiving older messages in conversation threads while preserving the **head message** (the most recent/relevant one)
2. Moving duplicate emails to a dedicated "Duplicates" folder (if the folder exists)

### Detection Logic

**Replied-to Detection:**
1. Scans email headers: `Message-ID`, `In-Reply-To`, and `References`
2. Identifies emails whose `Message-ID` is referenced by other emails in your inbox
3. If an email is referenced by another email, it is an older message in a thread and will be archived
4. The head message (not referenced by anything) stays in your inbox

**Duplicate Detection:**
1. Groups emails by their `Message-ID` header
2. If multiple emails share the same `Message-ID`, they are duplicates
3. Keeps the first copy, moves the rest to the "Duplicates" folder

### Benefits
- **Preserves conversation heads** - Only the latest/most relevant message remains
- **Handles forks** - Multiple replies to the same parent are all detected
- **Thread-aware** - Works with complex reply chains and forwards
- **Removes duplicates** - Cleans up duplicate deliveries

## Installation

1. Clone or download this repository
2. Open Microsoft Edge and navigate to `edge://extensions/`
3. Enable **Developer mode** (toggle in the sidebar)
4. Click **Load unpacked**
5. Select the `butler` folder
6. The Butler icon will appear in your toolbar

## Usage

1. **Open Outlook Web** at [outlook.office.com](https://outlook.office.com) or [outlook.cloud.microsoft](https://outlook.cloud.microsoft)
2. **Click the Butler icon** in your browser toolbar
3. Wait for the token status to show ready (captured automatically)
4. **Click "Dry Run"** to preview which emails would be archived and duplicates found
5. **Click "Archive All"** to:
   - Move duplicates to the "Duplicates" folder (if it exists)
   - Archive all replied-to emails

### Duplicate Handling

To enable duplicate moving:
1. Manually create a folder named **"Duplicates"** in your Outlook mailbox root
2. Run "Archive All" - duplicates will be moved to this folder before archiving

If the "Duplicates" folder does not exist, Butler will skip duplicate handling and only archive replied-to emails.

## Files

```
butler/
├── manifest.json       # Extension configuration (Manifest V3)
├── background.js       # Service worker - OWA API calls & token capture
├── popup.html          # Popup UI structure
├── popup.css           # Popup styling (dark theme)
├── popup.js            # Popup logic & button handlers
├── icons/              # Extension icons (PNG)
│   ├── icon16.png
│   ├── icon48.png
│   └── icon128.png
├── tools/              # Development utilities
│   ├── cli.js          # CLI tool for testing OWA APIs
│   └── token-server.js # Local server for token sharing
├── LICENSE             # MIT License
└── README.md           # This file
```

## Development / CLI Testing

For development and debugging, you can use the CLI tools:

### 1. Start the token server
```bash
node tools/token-server.js
```

### 2. Open Outlook Web in your browser
The extension will automatically push tokens to the token server.

### 3. Run the CLI archiver
```bash
# Dry run (preview only)
node tools/cli.js --auto

# Archive one email (test)
node tools/cli.js --auto --archive-one

# Archive all replied-to emails
node tools/cli.js --auto --archive-all
```

## Permissions

| Permission | Purpose |
|------------|---------|
| `webRequest` | Capture Bearer tokens from OWA requests |
| `storage` | Store captured tokens locally |
| `activeTab` | Popup interaction |
| `host_permissions` | Access OWA and Microsoft Graph APIs |

## Token Storage

- Tokens are stored locally in the browser
- Maximum 50 tokens stored
- Tokens expire after 24 hours

## Technical Details

### OWA APIs Used
- **FindItem** - Fetch inbox messages with headers
- **GetFolder** - Resolve distinguished folder IDs (e.g., `archive`)
- **FindFolder** - Search for folders by name (e.g., "Duplicates")
- **MoveItem** - Move emails to target folder

### Processing Order
1. Fetch up to 2000 emails from inbox
2. Detect duplicates by Message-ID
3. Move duplicates to "Duplicates" folder (if folder exists)
4. Re-fetch inbox after duplicate removal
5. Identify replied-to emails (emails whose Message-ID is referenced by others)
6. Move replied-to emails to Archive folder

## Compatibility

- Microsoft Edge (Chromium-based)
- Google Chrome
- Outlook Web App (outlook.office.com, outlook.cloud.microsoft)

## License

MIT - For personal use.
