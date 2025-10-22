# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

Outlook Email Exporter - A .NET 8.0 console application that exports emails from Microsoft 365/Outlook.com mailboxes to JSON format using Microsoft Graph API with Device Code Flow authentication.

## Quick Setup for SAMSYS Team Members

**For colleagues within the SAMSYS organization:**

Admin consent has already been granted organization-wide. You only need to:
1. Copy `appsettings.Example.json` to `appsettings.Development.json`
2. Set `ClientId`: `5723b5d0-bf95-4e8f-97f4-dbaf30a9fad9`
3. Set `TenantId`: `0b474a1c-e4d1-477f-95cb-9a74ddada3a3`
4. Run `dotnet restore && dotnet run`

**No Azure Portal work needed!** See README.md for full details.

## Common Commands

### Build and Run
```bash
# Build the project
dotnet build

# Run the application (interactive mode)
dotnet run --project OutlookExporter

# Run with command-line arguments
dotnet run --project OutlookExporter -- -m "mailbox@example.com" -f "Inbox"
dotnet run --project OutlookExporter -- --mailbox "mailbox@example.com" --folder "Sent Items"
dotnet run --project OutlookExporter -- -m "mailbox@example.com" -f "Inbox" -c 100  # Export 100 emails
dotnet run --project OutlookExporter -- -m "mailbox@example.com" -f "Inbox" -c 0    # Export all emails

# Export to HTML format
dotnet run --project OutlookExporter -- -m "mailbox@example.com" -f "Inbox" -o html

# Export to both JSON and HTML
dotnet run --project OutlookExporter -- -m "mailbox@example.com" -f "Inbox" -o both

# Show help
dotnet run --project OutlookExporter -- --help
```

### Development
```bash
# Restore dependencies
dotnet restore

# Clean build artifacts
dotnet clean
```

## Architecture

### Project Structure
- **OutlookExporter/Program.cs** - Main application logic (single-file console app)
- **OutlookExporter/OutlookExporter.csproj** - Project file with dependencies
- Configuration files are copied to output directory on build

### Key NuGet Dependencies
- Microsoft.Graph (v5.94.0) - Graph API client
- Azure.Identity (v1.17.0) - Device Code Flow authentication
- Microsoft.Extensions.Configuration.Json (v9.0.10) - Configuration management

### Application Flow
1. **Command-line argument parsing** (Program.cs:9-52) - Parses `-m/--mailbox`, `-f/--folder`, and `-c/--count` arguments
2. **Configuration loading** (Program.cs:69-79) - Loads appsettings.json and appsettings.Development.json
3. **Authentication** (Program.cs:63-91) - Device Code Flow with user browser authentication
4. **Mailbox discovery** (Program.cs:93-179) - Discovers available mailboxes (skipped if mailbox specified via CLI)
   - Primary mailbox
   - Known hardcoded mailbox (arquivo.comdev@samsys.pt)
   - Shared/delegated mailboxes via Azure AD query
5. **Mailbox selection** (Program.cs:181-274) - Interactive or CLI-based selection
6. **Folder retrieval** (Program.cs:276-423) - Recursive folder enumeration including subfolders
7. **Folder selection** (Program.cs:339-418) - Interactive or CLI-based selection
8. **Email export** (Program.cs:425-511) - Retrieves 5 emails and exports to JSON

### Authentication Architecture
- Uses Device Code Flow (OAuth 2.0) - no client secrets
- Supports both personal accounts (TenantId: "consumers") and organizational accounts (specific tenant ID)
- Required scopes: User.Read, User.Read.All, Mail.Read, Mail.ReadBasic, Mail.Read.Shared, MailboxSettings.Read
- Personal accounts work immediately; organizational accounts require admin consent

### Folder Enumeration
The application recursively retrieves all mail folders using a nested function pattern with full pagination support:
- `GetFoldersRecursive()` (Program.cs:319-358) - Recursively traverses folder hierarchy with pagination
- Builds folder paths as "Parent/Child/Grandchild"
- Root folders retrieved first (Program.cs:360-390), then children enumerated recursively
- Uses `PageIterator` to handle pagination and retrieve ALL folders (up to 999 per page)
- Critical for mailboxes with many folders - ensures no folders are missed

### Exported Email Data
JSON export includes all email properties except attachments:
- Identifiers (Id, InternetMessageId, ConversationId)
- Sender/Recipients (From, To, Cc, Bcc, ReplyTo)
- Dates (ReceivedDateTime, SentDateTime)
- Content (Subject, Body with HTML, BodyPreview)
- Metadata (IsRead, IsDraft, Importance, Categories, Flag)

## Configuration Requirements

### appsettings.json Structure
```json
{
  "AzureAd": {
    "ClientId": "your-client-id-here",
    "TenantId": "consumers"  // or specific tenant ID for organizational accounts
  },
  "KnownMailboxes": [
    {
      "DisplayName": "Example Shared Mailbox",
      "Email": "shared@example.com"
    }
  ]
}
```

### KnownMailboxes (Optional)
The `KnownMailboxes` configuration allows you to specify mailboxes that should always appear in the mailbox selection list. This is useful for:
- Delegated mailboxes that don't appear in automatic discovery
- Frequently accessed shared mailboxes
- Organization-specific mailboxes

Each entry requires:
- `DisplayName`: Friendly name shown in the mailbox list
- `Email`: Email address of the mailbox

These mailboxes are added to the list with type "Known" during mailbox discovery (Program.cs:123-139).

### Account Type Configuration
- **Personal accounts (Hotmail/Outlook.com)**: TenantId = "consumers"
- **Organizational accounts (Microsoft 365)**: TenantId = specific tenant ID or "common"

### Azure App Registration Requirements
- Application (client) ID configured in appsettings.json
- API permissions: Mail.Read, Mail.ReadBasic, MailboxSettings.Read, User.Read, User.Read.All, Mail.Read.Shared (all delegated)
- "Allow public client flows" must be enabled
- Organizational tenants require admin consent (see ADMIN_SETUP_GUIDE.md)

## Important Implementation Details

### Mailbox Discovery Behavior
Mailbox discovery (Program.cs:126-253) is ONLY executed when NO mailbox is specified via command-line argument. This prevents unnecessary API calls when the target mailbox is already known.

The discovery process includes:
1. Primary mailbox
2. Known mailboxes from configuration
3. Shared/delegated mailboxes (via Azure AD query)
4. **Archive mailboxes** - Automatically discovers Online Archive mailboxes for all accessible mailboxes (Program.cs:211-253)

Archive mailboxes are accessed using the naming pattern: `{localpart}-archive@{domain}` and appear in the mailbox list with "(Archive)" suffix.

### Folder Matching Logic
When a folder is specified via CLI argument, the application searches by both DisplayName and full Path (case-insensitive). If not found, it lists available folders and exits (Program.cs:358-387).

### Known Mailboxes Configuration
The application supports configuring known mailboxes via `appsettings.json` in the `KnownMailboxes` array. These mailboxes appear in the mailbox selection list with type "Known". This is useful for delegated mailboxes that don't appear in automatic discovery.

### Microsoft Graph API Rate Limits
The application is subject to Microsoft Graph API throttling limits:
- Be mindful when increasing the Top parameter (currently 5 emails, Program.cs:434)
- For bulk exports, implement pagination and retry logic
- Consider implementing exponential backoff for 429 responses

### Error Handling Patterns
- Folder not found: Lists available folders and exits (Program.cs:374-386)
- Mailbox access validation: Tests access and provides actionable error messages (Program.cs:226-253)
- Shared mailbox discovery: Silent failures for inaccessible mailboxes (Program.cs:141-160)

## Development Notes

### Single-File Architecture
The entire application logic is in Program.cs using top-level statements. This is intentional for simplicity. If adding significant features, consider refactoring into classes/services.

### Output Files
- Exported JSON files: `exported_emails_{folderName}.json` in OutlookExporter directory
- Folder names are sanitized for filesystem compatibility (Program.cs:501)

### Email Count and Pagination
The application supports configurable email export counts:
- Default: 5 emails (if no `-c/--count` argument provided)
- Specific count: Use `-c <number>` to export that many emails (e.g., `-c 100`)
- All emails: Use `-c 0` to export all emails with automatic pagination (Program.cs:450-503)

When exporting all emails (`-c 0`), the application:
- Uses maximum page size of 1000 emails per request
- Implements automatic pagination using PageIterator
- Shows progress updates every 1000 emails
- May take several minutes for large folders

**Performance Notes:**
- Specific count (e.g., `-c 100`): Single API request
- All emails in large folder: Multiple paginated requests, respects rate limits
