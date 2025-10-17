# Outlook Email Exporter

A C# console application to export emails from Outlook.com/Microsoft 365 to JSON format.

## Project Overview

**Purpose**: Export emails from Microsoft 365 Outlook account to JSON files with filtering capabilities.

**Features**:
- Connect to Outlook.com/Microsoft 365 via Microsoft Graph API
- Device Code Flow authentication (user-friendly, no secrets management)
- Filter emails by specific folders
- Export email data to JSON format (includes all properties except attachments)

## Technology Stack

- **.NET Version**: 8.0 (LTS)
- **Authentication**: Microsoft Identity Platform (Device Code Flow)
- **API**: Microsoft Graph API
- **Output Format**: JSON

## Setup Instructions

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

### Step 2: Project Setup

- Created .NET 8.0 console application
- Project location: `OutlookExporter/`

## Development Progress

### Completed Steps

1. ✅ **Azure App Registration**: Registered app in Azure Portal with required permissions
2. ✅ **Project Creation**: Created .NET 8.0 console application
3. ✅ **Documentation**: Created README.md with ongoing updates
4. ✅ **NuGet Packages Installed**:
   - `Microsoft.Graph` v5.94.0
   - `Azure.Identity` v1.17.0
   - `Microsoft.Extensions.Configuration.Json` v9.0.10
5. ✅ **Configuration System**:
   - Created `appsettings.json` and `appsettings.Development.json` for secure config storage
   - Added to `.gitignore` to prevent committing sensitive data
   - Created `appsettings.Example.json` as template for repository
6. ✅ **Device Code Authentication**: Implemented and tested successfully
7. ✅ **Mail Folder Listing**: Successfully retrieves and displays all mail folders with item counts
8. ✅ **Personal Account Support**: Configured for Outlook.com/Hotmail personal accounts using "consumers" tenant

### Completed Features

✅ **All core functionality implemented and tested:**
- Email folder listing with item counts
- Email retrieval from folders
- JSON export with complete email data (excluding attachments)

## Usage

### Running the Application

1. **Build the project:**
   ```bash
   dotnet build
   ```

2. **Run the application:**
   ```bash
   dotnet run
   ```

3. **Authentication Flow:**
   - The app will display a URL and code
   - Open the URL in your browser: `https://www.microsoft.com/link`
   - Enter the code when prompted
   - Sign in with your Microsoft account
   - Approve the requested permissions

4. **What the App Does:**
   - Authenticates with your Microsoft account
   - Lists all your mail folders with item counts
   - Exports 5 sample emails from your Inbox to `exported_emails.json`

### Output

The application creates an `exported_emails.json` file containing:
- Email ID and metadata
- Subject, sender, recipients (To/Cc/Bcc)
- Date received and sent
- Complete email body (HTML and plain text)
- Read/unread status, importance, categories
- Conversation and message IDs
- All properties except attachments

### Current Implementation

**Note:** The current version exports 5 sample emails from the Inbox folder. To customize:
- Edit `Program.cs` line 89 to change the number of emails
- Change `"Inbox"` to target different folders
- Modify the folder ID to export from specific folders listed in the output

## Configuration

### Setup Configuration Files

1. Copy `appsettings.Example.json` to `appsettings.Development.json`
2. Edit `appsettings.Development.json` and replace `YOUR_CLIENT_ID_HERE` with your Azure App Registration Client ID
3. **Important**: For personal Microsoft accounts (Hotmail/Outlook.com), set `TenantId` to `"consumers"`
4. For organizational accounts, use `"common"` or your specific tenant ID

**Example for personal accounts:**
```json
{
  "AzureAd": {
    "ClientId": "your-actual-client-id-here",
    "TenantId": "consumers"
  }
}
```

### Configuration Files

- `appsettings.json` - Production settings (git-ignored)
- `appsettings.Development.json` - Development settings (git-ignored)
- `appsettings.Example.json` - Template file (safe to commit)
- `.gitignore` - Ensures sensitive config files are not committed

## Project Structure

```
OutlookExporter/
├── OutlookExporter.csproj       # Project file with NuGet packages
├── Program.cs                   # Main entry point with authentication and folder listing
├── appsettings.json             # Production config (git-ignored)
├── appsettings.Development.json # Development config (git-ignored)
├── appsettings.Example.json     # Config template (safe to commit)
└── README.md                    # This documentation file
```

## Notes

- Authentication uses Device Code Flow (no client secrets needed)
- Exports include: subject, sender, recipients, date, body, and all metadata
- Attachments are NOT included in the export
- Filtering is done by folder selection

## Troubleshooting

### Error: "The mailbox is either inactive, soft-deleted, or is hosted on-premise"

**Solution**: This error occurs when using an external/guest account or wrong tenant configuration.

- For **personal Microsoft accounts** (Hotmail/Outlook.com): Set `TenantId` to `"consumers"` in your config
- For **organizational accounts**: Use `"common"` or your specific tenant ID
- Ensure you're logging in with the correct account type during device code authentication

### Authentication Issues

- If authentication fails, check that all required API permissions are added in Azure Portal
- Ensure "Allow public client flows" is enabled in the Azure app registration
- For personal accounts, always use `TenantId: "consumers"`

## Example Output

```
Outlook Email Exporter
======================

Initializing authentication...
Attempting to authenticate...

To sign in, use a web browser to open the page https://www.microsoft.com/link
and enter the code ABC123XYZ to authenticate.

Authentication successful!
Logged in as: Your Name
Email: yourname@hotmail.com

==================================================
Retrieving mail folders...
==================================================

Found 10 mail folders:

  - Inbox
    ID: [folder-id]
    Total Items: 4268
    Unread Items: 136

  - Sent Items
    ID: [folder-id]
    Total Items: 1532
    Unread Items: 0

[... more folders ...]

==================================================
Exporting emails from Inbox to JSON...
==================================================

Retrieved 5 emails
✓ Exported 5 emails to: exported_emails.json
  File size: 44.56 KB

Export completed successfully.

Done.
```

## Development Log

**2025-10-17**:
- Initial project setup and Azure app registration
- Installed NuGet packages: Microsoft.Graph, Azure.Identity, Configuration.Json
- Implemented secure configuration system with gitignore
- Implemented Device Code Flow authentication
- Implemented mail folder listing functionality
- Fixed tenant configuration for personal Microsoft accounts (consumers)
- Successfully tested authentication and folder retrieval
- Implemented email retrieval with full property support
- Implemented JSON export functionality
- **Project completed and fully functional**
