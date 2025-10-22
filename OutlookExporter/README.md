# Outlook Email Exporter

A C# console application to export emails from Outlook.com/Microsoft 365 to JSON format with support for multiple mailboxes, flexible filtering, and automated exports.

## Project Overview

**Purpose**: Export emails from Microsoft 365 Outlook accounts (personal and organizational) to JSON files with advanced filtering and automation capabilities.

**Key Features**:
- ✅ Connect to Outlook.com/Microsoft 365 via Microsoft Graph API
- ✅ Device Code Flow authentication (user-friendly, no secrets management)
- ✅ Multi-mailbox support (personal mailbox + shared/delegated mailboxes)
- ✅ Automatic shared mailbox discovery
- ✅ Recursive folder discovery (discovers ALL nested folders)
- ✅ Configurable email count (export specific number or all emails)
- ✅ Command-line arguments for automation and scripting
- ✅ Export to JSON format with complete email metadata (excluding attachments)
- ✅ Support for both interactive and automated modes

## Technology Stack

- **.NET Version**: 8.0 (LTS)
- **Authentication**: Microsoft Identity Platform (Device Code Flow)
- **API**: Microsoft Graph API
- **Output Format**: JSON

---

## 🚀 Quick Setup for Team Members (SAMSYS Organization)

**If you're a colleague within the SAMSYS organization, setup is VERY simple:**

### ✅ What You Need to Do:

1. **Install .NET 8.0 SDK** (if not already installed)
2. **Copy the example config file:**
   ```bash
   # In the OutlookExporter folder:
   cp appsettings.Example.json appsettings.Development.json
   ```
3. **Edit `appsettings.Development.json`** and replace `YOUR_CLIENT_ID_HERE` with:
   ```json
   {
     "AzureAd": {
       "ClientId": "5723b5d0-bf95-4e8f-97f4-dbaf30a9fad9",
       "TenantId": "0b474a1c-e4d1-477f-95cb-9a74ddada3a3"
     }
   }
   ```
4. **Run the application:**
   ```bash
   dotnet restore
   dotnet run
   ```

### ❌ What You DON'T Need to Do:

- **NO Azure Portal access needed**
- **NO app registration needed**
- **NO admin consent needed** (already granted organization-wide)
- **NO additional permissions needed**

**That's it!** Admin consent has already been granted for the entire organization. You just need the configuration file with the correct ClientId and TenantId, then authenticate with your organizational account when prompted.

---

## Setup Instructions (For New Organizations or Personal Use)

### Step 1: Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "App registrations"
3. Click "+ New registration"
   - **Name**: Outlook Email Exporter
   - **Supported account types**: Accounts in any organizational directory and personal Microsoft accounts
   - Click "Register"
4. Copy the **Application (client) ID** from the Overview page
5. Configure API Permissions:
   - Go to "API permissions"
   - Add "Microsoft Graph" → "Delegated permissions"
   - Add: `Mail.Read`, `Mail.ReadBasic`, `MailboxSettings.Read`
6. Enable Public Client Flow:
   - Go to "Authentication"
   - Under "Advanced settings" → "Allow public client flows" → Set to **Yes**
   - Click "Save"

### Step 2: Install Dependencies

```bash
dotnet add package Microsoft.Graph --version 5.94.0
dotnet add package Azure.Identity --version 1.17.0
dotnet add package Microsoft.Extensions.Configuration.Json --version 9.0.10
dotnet add package Microsoft.Extensions.Configuration.Binder --version 9.0.10
```

### Step 3: Configure Application

1. Copy `appsettings.Example.json` to `appsettings.Development.json`
2. Edit `appsettings.Development.json`:
   - Replace `YOUR_CLIENT_ID_HERE` with your Azure App Client ID
   - Set `TenantId`:
     - **Personal accounts** (Hotmail/Outlook.com): Use `"consumers"`
     - **Organizational accounts**: Use your tenant ID from Azure Portal
3. (Optional) Add known mailboxes to the `KnownMailboxes` array

## Development Progress

### Completed Features

#### Core Functionality
- ✅ Device Code Flow authentication (supports personal and organizational accounts)
- ✅ Automatic mailbox discovery (finds shared/delegated mailboxes you have access to)
- ✅ Recursive folder enumeration with pagination (discovers ALL nested folders, not just top-level)
- ✅ Configurable email export count (default: 5, specific count, or all emails with pagination)
- ✅ JSON export with complete email metadata (excluding attachments)

#### Advanced Features (October 2025)
- ✅ **Command-line arguments** for automation:
  - `-m, --mailbox <email>` - Specify mailbox
  - `-f, --folder <name>` - Specify folder to export
  - `-c, --count <number>` - Number of emails to export (0 = all)
  - `-h, --help` - Show help
- ✅ **Folder pagination fix**: Now discovers all folders (previously missed folders due to pagination bug)
- ✅ **Configuration-based known mailboxes**: Define mailboxes in appsettings.json
- ✅ **Early exit optimization**: Stops searching once target folder is found (35% faster)
- ✅ **PageIterator implementation**: Handles large datasets efficiently

#### Enhancements Log
- **October 21, 2025**: Added early exit optimization for folder search
- **October 20, 2025**: Fixed critical folder pagination bug (308 → 1,445 folders discovered)
- **October 20, 2025**: Implemented configurable email count with PageIterator
- **October 20, 2025**: Moved hardcoded mailboxes to configuration
- **October 17, 2025**: Initial implementation with authentication and basic export

## Usage

### Quick Start (Interactive Mode)

```bash
# Build the project
dotnet build

# Run interactively
dotnet run
```

**What happens:**
1. Device Code authentication flow (sign in via browser)
2. Discovers all accessible mailboxes (personal + shared/delegated)
3. Lists all mailboxes and prompts you to select one
4. Discovers all folders recursively (including deeply nested folders)
5. Lists folders and prompts you to select one
6. Exports 5 emails (default) from selected folder to JSON

### Automated Mode (Command-Line Arguments)

**Export specific number of emails:**
```bash
# Export 100 emails from Inbox of specific mailbox (JSON format)
dotnet run -- -m "user@company.com" -f "Inbox" -c 100

# Export 50 emails from "Sent Items" (JSON format)
dotnet run -- -m "shared@company.com" -f "Sent Items" -c 50
```

**Export all emails from a folder:**
```bash
# Export ALL emails from specific folder (pagination handles large datasets)
dotnet run -- -m "user@company.com" -f "Inbox" -c 0
```

**Export to HTML format:**
```bash
# Export emails to HTML (single styled document)
dotnet run -- -m "user@company.com" -f "Inbox" -o html

# Export all emails to HTML
dotnet run -- -m "user@company.com" -f "Inbox" -c 0 -o html
```

**Export to both JSON and HTML:**
```bash
# Export to both formats simultaneously
dotnet run -- -m "user@company.com" -f "Inbox" -o both

# Export all emails to both formats
dotnet run -- -m "user@company.com" -f "Inbox" -c 0 -o both
```

**Quick mailbox export:**
```bash
# Specify mailbox only (will prompt for folder)
dotnet run -- -m "user@company.com"

# Specify folder only (will prompt for mailbox)
dotnet run -- -f "Inbox"
```

**Show help:**
```bash
dotnet run -- --help
```

### Command-Line Arguments

| Argument | Short | Description | Example |
|----------|-------|-------------|---------|
| `--mailbox` | `-m` | Email address of mailbox to export from | `-m "user@company.com"` |
| `--folder` | `-f` | Folder name or path to export | `-f "Inbox"` or `-f "Inbox/Clients/A"` |
| `--count` | `-c` | Number of emails to export (0 = all) | `-c 100` or `-c 0` |
| `--format` | `-o` | Output format: json, html, or both (default: json) | `-o html` or `-o both` |
| `--help` | `-h` | Show help message | `-h` |

### Output Files

The application creates files named based on the folder and format:
- `exported_emails_Inbox.json` - JSON format for Inbox folder
- `exported_emails_Inbox.html` - HTML format for Inbox folder
- `exported_emails_SentItems.json` - For Sent Items (JSON)
- `exported_emails_Inbox01-CLIENTESV-ZXBSLOG.html` - For nested folders (HTML)

**JSON structure includes:**
- Email ID, subject, body (HTML/text)
- Sender, recipients (To/Cc/Bcc)
- Dates (received, sent)
- Read/unread status, importance, categories
- Conversation and internet message IDs
- All metadata except attachments

**HTML output includes:**
- Professional, styled single-page document
- Email cards with metadata tables
- Full email body (rendered HTML or plain text)
- Read/unread, importance, and draft badges
- Responsive design (mobile and desktop friendly)
- Print-friendly CSS

## Configuration

### Configuration File Structure

The application uses a layered configuration system:

**appsettings.Example.json** (Template - safe to commit):
```json
{
  "AzureAd": {
    "ClientId": "YOUR_CLIENT_ID_HERE",
    "TenantId": "common"
  },
  "KnownMailboxes": [
    {
      "DisplayName": "Example Shared Mailbox",
      "Email": "shared@example.com"
    }
  ],
  "_instructions": "Copy this file to appsettings.json and replace YOUR_CLIENT_ID_HERE",
  "_knownMailboxesInfo": "Optional list of mailboxes for discovery"
}
```

**appsettings.Development.json** (Your actual config - git-ignored):
```json
{
  "AzureAd": {
    "ClientId": "5723b5d0-bf95-4e8f-97f4-dbaf30a9fad9",
    "TenantId": "consumers"
  },
  "KnownMailboxes": [
    {
      "DisplayName": "Arquivo ComDev - SAMSYS",
      "Email": "arquivo.comdev@samsys.pt"
    }
  ]
}
```

### Configuration Options

**AzureAd Section:**
- `ClientId` - Your Azure App Registration Client ID (required)
- `TenantId` - Tenant for authentication:
  - `"consumers"` - Personal accounts (Hotmail, Outlook.com)
  - `"common"` - Accept any account type
  - `"<guid>"` - Specific organizational tenant ID

**KnownMailboxes Section (Optional):**
- Array of mailboxes to include in discovery
- Useful for delegated mailboxes not found by automatic discovery
- Each entry: `DisplayName` and `Email`

### Configuration Files

- `appsettings.json` - Base configuration (git-ignored)
- `appsettings.Development.json` - Development overrides (git-ignored)
- `appsettings.Example.json` - Template file (committed to git)
- `.gitignore` - Ensures sensitive configs are never committed

## Project Structure

```
outlook-export-2/
├── .gitignore                       # Prevents committing sensitive files
└── OutlookExporter/
    ├── OutlookExporter.csproj       # Project file with NuGet packages
    ├── Program.cs                   # Main application (~500 lines)
    ├── appsettings.json             # Base config (git-ignored)
    ├── appsettings.Development.json # Dev config (git-ignored)
    ├── appsettings.Example.json     # Config template (safe to commit)
    ├── README.md                    # This file
    ├── CLAUDE.md                    # Quick reference for Claude Code
    ├── PROJECT_SUMMARY.md           # Technical project summary
    ├── ADMIN_SETUP_GUIDE.md         # Guide for system administrators
    ├── LEARNING_PLAN.md             # Comprehensive learning guide
    └── LEARNING_NOTES.md            # Detailed learning notes and Q&A
```

## Architecture Highlights

**Single-File Design:**
- All logic in `Program.cs` (~500 lines)
- Simple, maintainable for learning purposes
- Top-level statements (modern C# 10+)

**Key Components:**
1. **Configuration**: Secure, layered configuration with .gitignore
2. **Authentication**: Device Code Flow with Azure.Identity
3. **Mailbox Discovery**: Automatic shared mailbox enumeration
4. **Folder Discovery**: Recursive with pagination (discovers 1,000+ folders)
5. **Email Export**: Configurable count with PageIterator for large datasets
6. **CLI Arguments**: Full automation support

## Performance

**Folder Discovery:**
- Discovers 1,445+ folders in ~3-4 seconds (with early exit optimization)
- Handles pagination automatically (999 items per page)

**Email Export:**
- Default: 5 emails (1 request)
- Specific count: Configurable (1-N requests depending on pagination)
- All emails: Uses PageIterator with progress indicators (handles thousands)

**Mailbox Discovery:**
- Interactive mode: 30-60 seconds (tests ~47 potential mailboxes)
- Automated mode: Skipped when mailbox specified via `-m` argument

## Notes

- Authentication uses Device Code Flow (no client secrets needed)
- Supports both personal and organizational Microsoft accounts
- Automatic shared mailbox discovery for organizational accounts
- Recursive folder discovery with pagination (finds ALL nested folders)
- Attachments are NOT included in exports
- Safe for automation and scheduling (Windows Task Scheduler compatible)

## Troubleshooting

### Common Issues

#### Error: "The mailbox is either inactive, soft-deleted, or is hosted on-premise"

**Cause**: Tenant ID mismatch between configuration and account type.

**Solution**:
- **Personal accounts** (Hotmail/Outlook.com): Set `TenantId` to `"consumers"`
- **Organizational accounts**: Use your specific tenant ID (GUID) from Azure Portal
- Ensure you're signing in with the correct account type

#### Error: "Needs administrator approval"

**Cause**: Organizational tenant requires admin consent for the application.

**Solution**:
1. Contact your IT administrator
2. Provide them with `ADMIN_SETUP_GUIDE.md`
3. Admin grants consent in Azure Portal → App Registration → API Permissions
4. Wait 5-10 minutes for propagation
5. Try again

#### Error: "Folder not found"

**Cause**: Folder name or path doesn't match exactly.

**Solution**:
- Run without `-f` argument to see list of all available folders
- Use exact folder name or full path (case-insensitive)
- For nested folders, use full path: `"Inbox/Clients/A/Company"`

#### Only discovering ~10-20 folders

**Cause**: This was a bug fixed in October 2025.

**Solution**:
- Ensure you have the latest version with folder pagination fix
- Should discover 1,000+ folders if they exist
- Check `LEARNING_NOTES.md` for details on the pagination bug fix

### Authentication Issues

- **Missing permissions**: Add all required permissions in Azure Portal (Mail.Read, User.Read, etc.)
- **Public client flows disabled**: Enable in Azure Portal → Authentication → Advanced Settings
- **Wrong tenant**: Personal accounts must use `"consumers"`, not organizational tenant ID

## Example Output

### Interactive Mode

```
Outlook Email Exporter
======================

Client ID: 5723b5d0-****
Tenant ID: consumers

Attempting to authenticate...

To sign in, use a web browser to open the page https://microsoft.com/devicelogin
and enter the code ABC123XYZ to authenticate.

Authentication successful!
Logged in as: Your Name
Email: yourname@hotmail.com

Adding 1 known mailbox(es) from configuration...

Attempting to discover shared/delegated mailboxes...
Testing 47 potential shared mailboxes for access...
Discovered 2 accessible shared mailbox(es).

Available mailboxes:
  [1] Your Name (yourname@hotmail.com) - Primary
  [2] Arquivo ComDev - SAMSYS (arquivo.comdev@samsys.pt) - Known
  [3] Shared Team Mailbox (team@company.com) - Shared

Select a mailbox (1-3): 2

Using mailbox: arquivo.comdev@samsys.pt

Discovering mail folders recursively...
Found 1,445 mail folders:
  [1] Inbox (4,268 items, 136 unread)
  [2] Sent Items (1,532 items, 0 unread)
  [3] Inbox/01-CLIENTES/A/Aber (23 items, 0 unread)
  [4] Inbox/01-CLIENTES/V-Z/XBSLOG (62 items, 5 unread)
  ... [1,441 more folders]

Select a folder (1-1445): 4

Exporting 5 email(s) from folder: Inbox/01-CLIENTES/V-Z/XBSLOG

Retrieved 5 emails
✓ Exported 5 emails to: exported_emails_Inbox01-CLIENTESV-ZXBSLOG.json
  File size: 45.23 KB

Export completed successfully.
```

### Automated Mode

```bash
$ dotnet run -- -m "arquivo.comdev@samsys.pt" -f "Inbox/01-CLIENTES/V-Z/XBSLOG" -c 0

Outlook Email Exporter
======================

Authentication successful!
Logged in as: Your Name

Skipping mailbox discovery (mailbox specified via command-line).
Using mailbox: arquivo.comdev@samsys.pt

Discovering mail folders recursively...
Found 932 mail folder(s) (stopped early - target folder found).

✓ Found folder: Inbox/01-CLIENTES/V-Z/XBSLOG

Exporting all emails from folder...
Retrieved 1000 emails...
Retrieved 2000 emails...
Total emails retrieved: 2,341

✓ Exported 2,341 emails to: exported_emails_Inbox01-CLIENTESV-ZXBSLOG.json
  File size: 1.87 MB

Export completed successfully.
```

## Use Cases

### Personal Use
- Backup personal email to JSON
- Export specific folders for archival
- Migrate email data between services
- Email analysis and reporting

### Organizational Use
- Export shared mailbox data for compliance
- Backup departmental mailboxes
- Email discovery and e-discovery
- Automated scheduled exports via Task Scheduler
- Mailbox auditing and reporting

### Automation
- Schedule exports with Windows Task Scheduler
- Integrate with backup systems
- Scripted batch exports of multiple mailboxes
- CI/CD pipeline email exports

## Future Enhancements

**Planned features** (see `LEARNING_PLAN.md` for details):
- Date range filtering
- Attachment download support
- Multiple output formats (CSV, Excel, EML)
- Advanced search and filtering
- Incremental exports (only new emails)
- Retry logic and throttling protection

## Documentation

- **README.md** (this file) - User guide and quick reference
- **CLAUDE.md** - Quick reference for Claude Code instances
- **PROJECT_SUMMARY.md** - Technical overview and architecture
- **ADMIN_SETUP_GUIDE.md** - Enterprise deployment guide
- **LEARNING_PLAN.md** - Comprehensive learning journey (from zero to production)
- **LEARNING_NOTES.md** - Detailed Q&A, concepts, and troubleshooting

## Contributing

This is a learning project. Key learnings are documented in `LEARNING_PLAN.md` and `LEARNING_NOTES.md`.

## License

This project is for educational purposes.

---

## Development Log

### October 21, 2025
- ✅ Implemented early exit optimization for folder search
- ✅ Performance improvement: 35% faster when searching for specific folders
- ✅ Added conditional output (show all folders vs report count)

### October 20, 2025
- ✅ **Critical Bug Fix**: Fixed folder pagination (308 → 1,445 folders discovered)
- ✅ Implemented PageIterator for root and child folders
- ✅ Added configurable email count feature (`-c` argument)
- ✅ Implemented PageIterator for unlimited email exports (`-c 0`)
- ✅ Moved hardcoded mailboxes to configuration
- ✅ Added `Microsoft.Extensions.Configuration.Binder` package
- ✅ Created `KnownMailboxes` configuration section

### October 17, 2025
- ✅ Initial project setup and Azure app registration
- ✅ Installed base NuGet packages (Microsoft.Graph, Azure.Identity, Configuration.Json)
- ✅ Implemented secure configuration system with .gitignore
- ✅ Implemented Device Code Flow authentication
- ✅ Implemented recursive folder discovery
- ✅ Fixed tenant configuration for personal accounts (`"consumers"`)
- ✅ Implemented email retrieval and JSON export
- ✅ Added multi-mailbox support
- ✅ Added shared mailbox discovery
- ✅ Added command-line arguments (`-m`, `-f`, `-h`)
- ✅ Created comprehensive documentation (5 markdown files)

**Status**: Production-ready with continuous enhancements
