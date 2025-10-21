# Outlook Email Exporter Learning Plan: From Zero to Production

## Project: "Outlook Email Exporter" - Your First Microsoft Graph API Application

---

## Table of Contents
1. [How This Learning Plan Works](#how-this-learning-plan-works)
2. [Project Overview](#project-overview)
3. [Key Concepts & Technologies](#key-concepts--technologies)
4. [Phase 1: Azure Setup & Configuration](#phase-1-azure-setup--configuration)
5. [Phase 2: Project Setup & Authentication](#phase-2-project-setup--authentication)
6. [Phase 3: Mail Access](#phase-3-mail-access)
7. [Phase 4: JSON Export](#phase-4-json-export)
8. [Phase 5: Multi-Mailbox Support](#phase-5-multi-mailbox-support)
9. [Project Structure](#project-structure)
10. [Testing & Debugging](#testing--debugging)
11. [Troubleshooting Guide](#troubleshooting-guide)
12. [Additional Resources](#additional-resources)

---

## How This Learning Plan Works

### Learning Methodology

This project follows a **guided, incremental learning approach**:

**📝 Very Small Steps**
- Each task broken into small, manageable pieces
- One feature at a time
- Validate each step before moving on
- Document all learnings and challenges

**🔍 Problem-Solving Focus**
- Real errors encountered and solved
- Understanding "why" not just "how"
- Learning from mistakes
- Building troubleshooting skills

**📚 Documentation-Driven Development**
- README tracks progress
- PROJECT_SUMMARY captures technical details
- ADMIN_SETUP_GUIDE for enterprise deployment
- All questions and solutions documented

**🧪 Test Early, Test Often**
- Build and run after each significant change
- Verify authentication before API calls
- Test with different account types
- Document what works and what doesn't

### Example Development Flow

```
1. Azure setup → Test login
2. Add NuGet packages → Build project
3. Implement auth → Test device code flow
4. Add folder listing → Verify API access
5. Add JSON export → Check output file
6. Add mailbox selection → Test with shared mailboxes
```

### Your Responsibilities

- ✅ Follow steps in order
- ✅ Test after each major change
- ✅ Ask questions when concepts are unclear
- ✅ Document errors encountered
- ✅ Don't skip configuration steps

### What You'll Learn

- ✅ Azure Active Directory and App Registration
- ✅ OAuth 2.0 and Device Code Flow
- ✅ Microsoft Graph API usage
- ✅ C# async/await patterns
- ✅ Secure configuration management
- ✅ Enterprise authentication (admin consent)
- ✅ API permission models
- ✅ JSON serialization in .NET

---

## Project Overview

### What You'll Build

A production-ready C# console application that:
- Authenticates with Microsoft 365/Outlook accounts
- Lists mail folders with item counts
- Exports emails to JSON format
- Supports both personal (Hotmail/Outlook.com) and organizational (Microsoft 365) accounts
- Accesses shared/delegated mailboxes
- Includes all email metadata except attachments
- Uses secure configuration practices
- Works with Device Code Flow (no password storage)

### Final Result

By the end of this journey, you'll have:
- A working console app that exports emails
- Understanding of Azure AD and OAuth 2.0
- Knowledge of Microsoft Graph API
- Experience with tenant IDs and admin consent
- Skills to build enterprise-ready authentication
- Foundation for building other Microsoft 365 integrations

### Real-World Use Cases

- Email backup and archival
- Data migration between systems
- Email analysis and reporting
- Compliance and e-discovery
- Mailbox auditing
- Custom email processing workflows

---

## Key Concepts & Technologies

### 1. Azure Active Directory (Azure AD)

**What is it?**
Azure AD is Microsoft's cloud-based identity and access management service.

**Key Points:**
- Manages users, groups, and applications
- Handles authentication (who you are)
- Handles authorization (what you can access)
- Required for accessing Microsoft 365 services
- Different from Active Directory (on-premise)

**For This Project:**
- You register your app in Azure AD
- Azure AD handles user authentication
- Azure AD enforces permissions
- No password storage in your app!

### 2. App Registration

**What is it?**
Registering your application in Azure Portal to get identity and permissions.

**What You Get:**
- **Client ID**: Your app's unique identifier
- **Tenant ID**: Your organization's identifier
- **Permissions**: What your app can access

**Why Needed?**
- Microsoft needs to know who's requesting access
- Security: prevents unauthorized API access
- Audit trail: track which apps access data
- User consent: users approve what app can do

### 3. OAuth 2.0 & Device Code Flow

**What is OAuth 2.0?**
An authorization framework that lets apps access resources without handling passwords.

**Device Code Flow:**
A specific OAuth 2.0 flow designed for devices with limited input capabilities.

**How It Works:**
```
1. App requests device code from Azure AD
2. Azure AD returns:
   - User code (e.g., "ABC123")
   - Verification URL (e.g., microsoft.com/devicelogin)
3. App displays: "Go to URL and enter code"
4. User opens browser, goes to URL, enters code
5. User signs in and consents
6. App polls Azure AD for token
7. Azure AD returns access token
8. App uses token to call Microsoft Graph API
```

**Why Use Device Code Flow?**
- ✅ No browser control needed (works in console apps)
- ✅ No redirect URI complexity
- ✅ No password handling
- ✅ User-friendly (browser-based sign-in)
- ✅ Works on any device with browser access

### 4. Microsoft Graph API

**What is it?**
A unified REST API endpoint for accessing Microsoft 365 services.

**Endpoints:**
- `https://graph.microsoft.com/v1.0/` - Production endpoint
- Accesses: Mail, Calendar, Contacts, OneDrive, Teams, etc.

**Key Concepts:**
- **Resources**: Things you access (users, messages, folders)
- **Methods**: HTTP verbs (GET, POST, PATCH, DELETE)
- **Permissions**: What your app can do

**For This Project:**
```
GET /me                           → Get signed-in user profile
GET /me/mailFolders               → List mail folders
GET /me/mailFolders/{id}/messages → Get emails from folder
GET /users/{email}/mailFolders    → Access other mailboxes
```

### 5. Permissions: Delegated vs Application

**Delegated Permissions:**
- App acts on behalf of a signed-in user
- User must be present
- App can only access what user can access
- Examples: Mail.Read, User.Read

**Application Permissions:**
- App acts on its own (no user present)
- Requires admin consent
- Used for background services/daemons
- Examples: Mail.Read.All (all mailboxes)

**This Project Uses:** Delegated permissions (user context)

### 6. Tenant IDs

**What is a Tenant?**
An instance of Azure AD representing an organization.

**Critical Concept:**

| Account Type | Tenant ID | Why |
|--------------|-----------|-----|
| **Personal** (Hotmail, Outlook.com) | `"consumers"` | Special tenant for Microsoft personal accounts |
| **Organizational** (Microsoft 365) | Specific GUID (e.g., `0b474a1c-...`) | Your organization's tenant |
| **Multi-tenant** | `"common"` | Accepts any account type |

**Common Error:**
Using organizational tenant ID with personal account → "Mailbox inactive" error

### 7. Admin Consent

**What is it?**
Administrator approval for an app to access organizational data.

**Why Needed for Organizational Accounts?**
- Security: prevents unauthorized apps
- Governance: IT controls what apps can run
- Compliance: audit trail of approved apps

**Personal Accounts:** No admin consent needed (user can consent directly)

**Organizational Accounts:** Admin must grant consent in Azure Portal

### 8. Async/Await in C#

**What is it?**
C# pattern for asynchronous programming.

**Key Points:**
```csharp
// Synchronous (blocks thread)
var user = GetUser();

// Asynchronous (non-blocking)
var user = await GetUserAsync();
```

**Why Use It?**
- Network calls take time
- Don't block the application
- Better performance and responsiveness

**In This Project:**
All Microsoft Graph API calls are async:
```csharp
await graphClient.Me.GetAsync();
await graphClient.Me.MailFolders.GetAsync();
await graphClient.Me.MailFolders["Inbox"].Messages.GetAsync();
```

---

## Phase 1: Azure Setup & Configuration (1-2 hours)

### Goal
Set up Azure App Registration with proper permissions and configuration.

### Step 1.1: Create Azure App Registration

**What you'll learn**: Azure Portal navigation, app registration basics

**Process**:
1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory"
3. Select "App registrations" → "New registration"
4. Enter name: "Outlook Email Exporter"
5. Select account types: "Accounts in any organizational directory and personal Microsoft accounts"
6. No redirect URI needed (Device Code Flow)
7. Click "Register"

**What You Get**:
- Application (client) ID
- Directory (tenant) ID

**Save These**: You'll need them for configuration

**Testing**:
- Verify app appears in "App registrations" list
- Copy Client ID and Tenant ID

### Step 1.2: Configure API Permissions

**What you'll learn**: Permission types, scopes, Microsoft Graph permissions

**Process**:
1. In your app registration, click "API permissions"
2. Click "+ Add a permission"
3. Select "Microsoft Graph"
4. Select "Delegated permissions"
5. Search for and add:
   - `User.Read` - Read user profile
   - `Mail.Read` - Read user mail
   - `Mail.ReadBasic` - Basic mail access
   - `MailboxSettings.Read` - Read mailbox settings
6. Click "Add permissions"

**Understanding Permissions**:
- `User.Read`: Get signed-in user info (name, email)
- `Mail.Read`: Read emails from mailbox
- `Mail.ReadBasic`: Read basic email properties
- `MailboxSettings.Read`: Access mailbox configuration

**Testing**:
- Verify all 4 permissions appear in list
- Check they're marked as "Delegated"

### Step 1.3: Enable Public Client Flow

**What you'll learn**: Public vs confidential clients, Device Code Flow requirements

**Process**:
1. In app registration, click "Authentication"
2. Scroll to "Advanced settings"
3. Under "Allow public client flows", set to **Yes**
4. Click "Save"

**Why?**
- Device Code Flow requires public client flows enabled
- Console apps are "public clients" (can't securely store secrets)
- Without this: authentication will fail

**Testing**:
- Verify toggle is set to "Yes"
- Save setting persists after refresh

### Step 1.4: Understand Tenant IDs

**What you'll learn**: Personal vs organizational accounts, tenant selection

**Key Decision Point**:

**For Personal Accounts (Hotmail/Outlook.com):**
- Use Tenant ID: `"consumers"`
- No admin consent needed
- Immediate authentication

**For Organizational Accounts (Microsoft 365):**
- Use your specific Tenant ID (from Overview page)
- May require admin consent
- Check with IT department

**Testing**:
- Know which account type you'll test with
- Have correct Tenant ID ready

---

## Phase 2: Project Setup & Authentication (1-2 hours)

### Goal
Create .NET project, add packages, implement secure configuration, and authenticate successfully.

### Step 2.1: Create .NET Console Project

**What you'll learn**: .NET CLI, project structure

**Commands**:
```bash
dotnet new console -n OutlookExporter
cd OutlookExporter
```

**Testing**:
```bash
dotnet build
dotnet run
```
Should see "Hello, World!"

### Step 2.2: Install NuGet Packages

**What you'll learn**: NuGet package management, dependencies

**Packages to Install**:
```bash
dotnet add package Microsoft.Graph --version 5.94.0
dotnet add package Azure.Identity --version 1.17.0
dotnet add package Microsoft.Extensions.Configuration.Json --version 9.0.10
```

**What Each Does**:
- `Microsoft.Graph`: Microsoft Graph API client library
- `Azure.Identity`: Authentication libraries (DeviceCodeCredential)
- `Microsoft.Extensions.Configuration.Json`: Read JSON configuration files

**Testing**:
```bash
dotnet build
```
Should build successfully with no errors.

### Step 2.3: Create Configuration System

**What you'll learn**: Secure configuration, .gitignore, JSON structure

**Files to Create**:

**1. appsettings.Example.json** (safe to commit):
```json
{
  "AzureAd": {
    "ClientId": "YOUR_CLIENT_ID_HERE",
    "TenantId": "consumers"
  },
  "_instructions": "Copy to appsettings.json and replace placeholders"
}
```

**2. appsettings.Development.json** (git-ignored):
```json
{
  "AzureAd": {
    "ClientId": "your-actual-client-id",
    "TenantId": "consumers"
  }
}
```

**3. .gitignore** (prevent credential leaks):
```
appsettings.json
appsettings.Development.json
bin/
obj/
```

**Why This Approach?**
- ✅ No hardcoded credentials in code
- ✅ Safe template for version control
- ✅ Environment-specific configs
- ✅ Prevents accidental credential commits

**Testing**:
- Verify .gitignore prevents committing sensitive files
- Check appsettings.Example.json is tracked by git
- Verify appsettings.Development.json is ignored

### Step 2.4: Implement Configuration Loading

**What you'll learn**: IConfiguration, configuration builder

**Code in Program.cs**:
```csharp
using Microsoft.Extensions.Configuration;

var configuration = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddJsonFile("appsettings.Development.json", optional: true, reloadOnChange: true)
    .Build();

var clientId = configuration["AzureAd:ClientId"]
    ?? throw new InvalidOperationException("ClientId not found in configuration");
var tenantId = configuration["AzureAd:TenantId"] ?? "common";

Console.WriteLine($"Client ID: {clientId}");
Console.WriteLine($"Tenant ID: {tenantId}");
```

**Testing**:
```bash
dotnet run
```
Should print your Client ID and Tenant ID.

### Step 2.5: Implement Device Code Authentication

**What you'll learn**: DeviceCodeCredential, async/await, GraphServiceClient

**Code in Program.cs**:
```csharp
using Azure.Identity;
using Microsoft.Graph;

var scopes = new[] { "User.Read", "Mail.Read", "Mail.ReadBasic", "MailboxSettings.Read" };

var options = new DeviceCodeCredentialOptions
{
    ClientId = clientId,
    TenantId = tenantId,
    DeviceCodeCallback = (code, cancellation) =>
    {
        Console.WriteLine("\n" + code.Message);
        return Task.CompletedTask;
    }
};

var credential = new DeviceCodeCredential(options);
var graphClient = new GraphServiceClient(credential, scopes);

Console.WriteLine("\nAttempting to authenticate...");
var user = await graphClient.Me.GetAsync();

Console.WriteLine($"\nAuthentication successful!");
Console.WriteLine($"Logged in as: {user?.DisplayName}");
Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName}");
```

**What This Does**:
1. Creates device code credential with callback
2. Displays device code URL and code to user
3. User opens browser, signs in, enters code
4. App polls for token
5. Gets user profile to verify authentication
6. Prints user name and email

**Testing**:
```bash
dotnet run
```
1. App displays URL and code
2. Open URL in browser
3. Enter code
4. Sign in with Microsoft account
5. Approve permissions
6. App prints your name and email

**Success Criteria**:
- ✅ Device code displayed
- ✅ Browser authentication works
- ✅ User info printed correctly
- ✅ No errors in console

---

## Phase 3: Mail Access (1-2 hours)

### Goal
Access mailbox, list folders, and troubleshoot common errors.

### Step 3.1: List Mail Folders

**What you'll learn**: Microsoft Graph API structure, mailbox resources

**Code to Add**:
```csharp
Console.WriteLine("\n" + new string('=', 50));
Console.WriteLine("Retrieving mail folders...");
Console.WriteLine(new string('=', 50));

var folders = await graphClient.Me.MailFolders.GetAsync();

if (folders?.Value != null && folders.Value.Count > 0)
{
    Console.WriteLine($"\nFound {folders.Value.Count} mail folders:\n");

    foreach (var folder in folders.Value)
    {
        Console.WriteLine($"  - {folder.DisplayName}");
        Console.WriteLine($"    ID: {folder.Id}");
        Console.WriteLine($"    Total Items: {folder.TotalItemCount}");
        Console.WriteLine($"    Unread Items: {folder.UnreadItemCount}");
        Console.WriteLine();
    }
}
else
{
    Console.WriteLine("\nNo folders found.");
}
```

**Understanding the API**:
- `graphClient.Me` = signed-in user's context
- `.MailFolders` = collection of mail folders
- `.GetAsync()` = asynchronous GET request
- Returns collection with `.Value` property

**Testing**:
```bash
dotnet run
```
Should display all your mail folders with counts.

### Step 3.2: Troubleshoot "Mailbox Inactive" Error

**If You Get Error**: "The mailbox is either inactive, soft-deleted, or is hosted on-premise"

**Root Cause**: Tenant ID mismatch

**Solution**:

**For Personal Accounts (Hotmail/Outlook.com)**:
Change `appsettings.Development.json`:
```json
{
  "AzureAd": {
    "ClientId": "your-client-id",
    "TenantId": "consumers"  ← Change to "consumers"
  }
}
```

**For Organizational Accounts (Microsoft 365)**:
Use your organization's specific tenant ID:
```json
{
  "AzureAd": {
    "ClientId": "your-client-id",
    "TenantId": "0b474a1c-e4d1-477f-95cb-9a74ddada3a3"  ← Your tenant ID
  }
}
```

**Why This Happens**:
- Personal accounts don't exist in organizational tenants
- Must use `"consumers"` tenant for personal accounts
- Must use specific tenant ID for organizational accounts

**Testing**:
After changing tenant ID:
```bash
dotnet run
```
Should now successfully list folders.

### Step 3.3: Troubleshoot "Admin Approval Required" Error

**If You Get Error** (organizational accounts): "Necessita de aprovação do administrador" or "Needs administrator approval"

**Root Cause**: Organizational tenants require admin consent

**Solution**:

**Option 1: Request Admin Consent**
1. Contact your IT administrator
2. Provide them with `ADMIN_SETUP_GUIDE.md`
3. Admin grants consent in Azure Portal
4. Wait 5-10 minutes for propagation
5. Try again

**Option 2: Admin Self-Service** (if you're the admin):
1. Go to Azure Portal
2. Navigate to your app registration
3. Click "API permissions"
4. Click "Grant admin consent for [Organization]"
5. Confirm

**Why This Happens**:
- Organizational policies require admin approval
- Prevents unauthorized apps from accessing company data
- Standard enterprise security practice

**Testing**:
After admin consent:
```bash
dotnet run
```
Should authenticate and list folders successfully.

---

## Phase 4: JSON Export (1 hour)

### Goal
Export emails to JSON with all metadata except attachments.

### Step 4.1: Retrieve Emails from Inbox

**What you'll learn**: Querying messages, request parameters

**Code to Add**:
```csharp
Console.WriteLine("\n" + new string('=', 50));
Console.WriteLine("Exporting emails from Inbox to JSON...");
Console.WriteLine(new string('=', 50));

var messages = await graphClient.Me.MailFolders["Inbox"].Messages
    .GetAsync(requestConfig =>
    {
        requestConfig.QueryParameters.Top = 5; // Limit for testing
    });

if (messages?.Value != null && messages.Value.Count > 0)
{
    Console.WriteLine($"\nRetrieved {messages.Value.Count} emails");
}
else
{
    Console.WriteLine("\nNo emails found in Inbox.");
}
```

**Understanding the API**:
- `.MailFolders["Inbox"]` = specific folder
- `.Messages` = collection of emails
- `requestConfig.QueryParameters.Top = 5` = limit results

**Testing**:
```bash
dotnet run
```
Should print number of retrieved emails.

### Step 4.2: Map Email Properties to Anonymous Objects

**What you'll learn**: LINQ, anonymous types, null-conditional operators

**Code to Add**:
```csharp
var emailData = messages.Value.Select(msg => new
{
    Id = msg.Id,
    Subject = msg.Subject,
    From = new
    {
        Name = msg.From?.EmailAddress?.Name,
        Address = msg.From?.EmailAddress?.Address
    },
    ToRecipients = msg.ToRecipients?.Select(r => new
    {
        Name = r.EmailAddress?.Name,
        Address = r.EmailAddress?.Address
    }).ToList(),
    CcRecipients = msg.CcRecipients?.Select(r => new
    {
        Name = r.EmailAddress?.Name,
        Address = r.EmailAddress?.Address
    }).ToList(),
    BccRecipients = msg.BccRecipients?.Select(r => new
    {
        Name = r.EmailAddress?.Name,
        Address = r.EmailAddress?.Address
    }).ToList(),
    ReceivedDateTime = msg.ReceivedDateTime,
    SentDateTime = msg.SentDateTime,
    HasAttachments = msg.HasAttachments,
    Importance = msg.Importance?.ToString(),
    IsRead = msg.IsRead,
    IsDraft = msg.IsDraft,
    InternetMessageId = msg.InternetMessageId,
    ConversationId = msg.ConversationId,
    Categories = msg.Categories,
    Body = new
    {
        ContentType = msg.Body?.ContentType?.ToString(),
        Content = msg.Body?.Content
    },
    BodyPreview = msg.BodyPreview
}).ToList();
```

**What This Does**:
- Creates anonymous objects (no class definition needed)
- Maps all email properties
- Flattens nested structures (From, ToRecipients)
- Handles nulls gracefully with `?.` operator

**Key Concepts**:
- `msg => new { ... }`: Lambda expression creating object
- `?.`: Null-conditional operator (safe navigation)
- `.Select()`: LINQ projection (transform each item)
- `.ToList()`: Materialize query

### Step 4.3: Serialize to JSON and Save

**What you'll learn**: JSON serialization, file I/O

**Code to Add**:
```csharp
using System.Text.Json;

var jsonOptions = new JsonSerializerOptions
{
    WriteIndented = true,
    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
};

var json = JsonSerializer.Serialize(emailData, jsonOptions);

var outputFile = "exported_emails.json";
await File.WriteAllTextAsync(outputFile, json);

Console.WriteLine($"✓ Exported {emailData.Count} emails to: {outputFile}");
Console.WriteLine($"  File size: {new FileInfo(outputFile).Length / 1024.0:F2} KB");
```

**JSON Options Explained**:
- `WriteIndented = true`: Pretty-print (readable formatting)
- `Encoder = UnsafeRelaxedJsonEscaping`: Handle special characters properly

**Testing**:
```bash
dotnet run
```
1. Should create `exported_emails.json`
2. Open file and verify structure
3. Check all properties are present
4. Verify JSON is valid and readable

**Success Criteria**:
- ✅ File created
- ✅ Valid JSON format
- ✅ All email properties present
- ✅ Readable formatting

---

## Phase 5: Multi-Mailbox Support & Advanced Features (2-3 hours)

### Goal
Access shared/delegated mailboxes, implement command-line automation, and add recursive folder discovery.

### Step 5.1: Understand Mailbox Access Patterns

**What you'll learn**: Me vs Users endpoint, shared mailbox permissions

**Primary Mailbox Access**:
```csharp
// Accesses signed-in user's mailbox
await graphClient.Me.MailFolders.GetAsync();
await graphClient.Me.MailFolders["Inbox"].Messages.GetAsync();
```

**Shared Mailbox Access**:
```csharp
// Accesses specific mailbox by email
await graphClient.Users["shared@company.com"].MailFolders.GetAsync();
await graphClient.Users["shared@company.com"].MailFolders["Inbox"].Messages.GetAsync();
```

**Requirements for Shared Mailbox**:
- User must have "Full Access" permission
- Or mailbox must be delegated to user
- Configured by Exchange admin

### Step 5.2: Implement Automatic Mailbox Discovery

**What you'll learn**: Azure AD queries, disabled accounts, shared mailbox discovery

**Add Permission**:
```csharp
var scopes = new[] {
    "User.Read",
    "User.Read.All",  // ← Add this for mailbox discovery
    "Mail.Read",
    "Mail.ReadBasic",
    "Mail.Read.Shared",
    "MailboxSettings.Read"
};
```

**Discovery Code**:
```csharp
Console.WriteLine("\nAttempting to discover shared/delegated mailboxes...");

// Query for disabled accounts (traditional shared mailboxes)
var users = await graphClient.Users
    .GetAsync(requestConfig =>
    {
        requestConfig.QueryParameters.Filter = "accountEnabled eq false";
        requestConfig.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName" };
    });

// Test access to each potential mailbox
foreach (var potentialMailbox in users.Value)
{
    try
    {
        await graphClient.Users[potentialMailbox.Mail]
            .MailFolders
            .GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Top = 1;
            });

        // Success! User has access
        availableMailboxes.Add((
            potentialMailbox.DisplayName ?? "Unknown",
            potentialMailbox.Mail ?? "",
            "Shared"
        ));
    }
    catch
    {
        // No access, skip
    }
}
```

**Key Learning**:
- Shared mailboxes often have `accountEnabled = false`
- Must test access to verify permissions
- Some mailboxes may have `accountEnabled = true` (delegated)

### Step 5.3: Implement Command-Line Arguments

**What you'll learn**: Argument parsing, automation, non-interactive mode

**Add Argument Parser**:
```csharp
// Parse command-line arguments
string? argMailbox = null;
string? argFolder = null;

for (int i = 0; i < args.Length; i++)
{
    if ((args[i] == "--mailbox" || args[i] == "-m") && i + 1 < args.Length)
    {
        argMailbox = args[i + 1];
        i++; // Skip next argument
    }
    else if ((args[i] == "--folder" || args[i] == "-f") && i + 1 < args.Length)
    {
        argFolder = args[i + 1];
        i++; // Skip next argument
    }
    else if (args[i] == "--help" || args[i] == "-h")
    {
        Console.WriteLine("Usage: OutlookExporter [options]");
        Console.WriteLine("\nOptions:");
        Console.WriteLine("  -m, --mailbox <email>    Specify mailbox email address");
        Console.WriteLine("  -f, --folder <name>      Specify folder name to export");
        Console.WriteLine("  -h, --help               Show this help message");
        Console.WriteLine("\nExamples:");
        Console.WriteLine("  OutlookExporter --mailbox user@example.com --folder \"Sent Items\"");
        Console.WriteLine("  OutlookExporter -m user@example.com -f Inbox");
        return;
    }
}
```

**Update Discovery Logic**:
```csharp
// Only discover mailboxes if not specified via command-line argument
if (argMailbox == null)
{
    // ... [mailbox discovery logic]
}
else
{
    Console.WriteLine("\nSkipping mailbox discovery (mailbox specified via command-line).");
    selectedMailboxEmail = argMailbox;
}
```

**Testing**:
```bash
# Interactive mode
dotnet run

# Automated mode
dotnet run -- -m "shared@company.com" -f "Sent Items"
```

**Benefits**:
- ✅ Scriptable and automatable
- ✅ No user interaction needed
- ✅ Faster execution (skips discovery)
- ✅ Can be scheduled via Task Scheduler

### Step 5.4: Implement Recursive Folder Discovery

**What you'll learn**: Recursive async functions, ChildFolders endpoint, folder hierarchies

**Problem**: Default folder listing only shows top-level folders (e.g., 10 folders)

**Solution**: Recursively discover all subfolders

**Recursive Discovery Code**:
```csharp
var allFolders = new List<(string Id, string Name, string Path, int Total, int Unread)>();

// Get top-level folders
var topLevelFolders = await graphClient.Users[selectedMailboxEmail]
    .MailFolders
    .GetAsync();

if (topLevelFolders?.Value != null)
{
    foreach (var folder in topLevelFolders.Value)
    {
        allFolders.Add((
            folder.Id ?? "",
            folder.DisplayName ?? "",
            folder.DisplayName ?? "",
            folder.TotalItemCount ?? 0,
            folder.UnreadItemCount ?? 0
        ));

        // Recursively get child folders
        if (folder.ChildFolderCount > 0)
        {
            await GetFoldersRecursive(folder.Id ?? "", folder.DisplayName ?? "");
        }
    }
}

async Task GetFoldersRecursive(string parentFolderId, string parentPath)
{
    var childFolders = await graphClient.Users[selectedMailboxEmail]
        .MailFolders[parentFolderId]
        .ChildFolders
        .GetAsync();

    if (childFolders?.Value != null)
    {
        foreach (var folder in childFolders.Value)
        {
            var folderPath = string.IsNullOrEmpty(parentPath)
                ? folder.DisplayName ?? ""
                : $"{parentPath}/{folder.DisplayName}";

            allFolders.Add((
                folder.Id ?? "",
                folder.DisplayName ?? "",
                folderPath,
                folder.TotalItemCount ?? 0,
                folder.UnreadItemCount ?? 0
            ));

            // Recursively get child folders
            if (folder.ChildFolderCount > 0)
            {
                await GetFoldersRecursive(folder.Id ?? "", folderPath);
            }
        }
    }
}
```

**Result**: Discovers ALL folders, including deeply nested ones (e.g., `Inbox/01-CLIENTES/A/Aber`)

**Testing**:
- Look for nested folder in output
- Verify folder paths show hierarchy
- Count should increase significantly (e.g., 10 → 308 folders)

### Step 5.5: Improve Error Handling

**What you'll learn**: User feedback, error messages, program exit

**Problem**: When folder not found, app defaults to Inbox (confusing)

**Solution**: Exit with helpful error message

**Error Handling Code**:
```csharp
var selectedFolder = allFolders.FirstOrDefault(f =>
    f.Name.Equals(argFolder, StringComparison.OrdinalIgnoreCase) ||
    f.Path.Equals(argFolder, StringComparison.OrdinalIgnoreCase)
);

if (selectedFolder.Id != null)
{
    Console.WriteLine($"✓ Found folder: {selectedFolder.Path}");
    selectedFolderId = selectedFolder.Id;
    selectedFolderName = selectedFolder.Name;
}
else
{
    Console.WriteLine($"✗ Error: Folder '{argFolder}' not found.");
    Console.WriteLine("\nAvailable folders:");
    foreach (var folder in allFolders.Take(10))
    {
        Console.WriteLine($"  - {folder.Path}");
    }
    if (allFolders.Count > 10)
    {
        Console.WriteLine($"  ... and {allFolders.Count - 10} more folders");
    }
    Console.WriteLine("\nPlease specify a valid folder name or path.");
    return;  // Exit program
}
```

**Benefits**:
- ✅ Clear error message
- ✅ Shows available folders
- ✅ Prevents incorrect exports
- ✅ Better user experience

### Step 5.6: Performance Optimization

**What you'll learn**: Conditional execution, performance tuning

**Problem**: Mailbox discovery takes 30-60 seconds (tests 47 mailboxes)

**Solution**: Skip discovery when mailbox specified via args

**Already Implemented in Step 5.3**:
```csharp
if (argMailbox == null)
{
    // Discovery logic here (30-60 seconds)
}
else
{
    // Skip discovery (instant)
}
```

**Performance Comparison**:
- Interactive mode: 30-60 seconds (discovers all mailboxes)
- Automated mode with args: Instant (skips discovery)

**Testing**:
```bash
# Slow (with discovery)
dotnet run

# Fast (skip discovery)
dotnet run -- -m "user@example.com" -f "Inbox"
```

---

## Project Structure

```
outlook-export-2/
├── .gitignore                       # Prevents committing sensitive files
└── OutlookExporter/
    ├── OutlookExporter.csproj       # Project file with NuGet packages
    ├── Program.cs                   # Main application logic (~240 lines)
    ├── appsettings.json             # Production config (git-ignored)
    ├── appsettings.Development.json # Development config (git-ignored)
    ├── appsettings.Example.json     # Config template (safe to commit)
    ├── README.md                    # User documentation
    ├── PROJECT_SUMMARY.md           # Technical project summary
    ├── ADMIN_SETUP_GUIDE.md         # Guide for system administrators
    ├── LEARNING_PLAN.md             # This file!
    ├── LEARNING_NOTES.md            # Detailed learning notes
    └── exported_emails.json         # Output file (generated at runtime)
```

---

## Rate Limits & Throttling

### Understanding Microsoft Graph API Limits

**What you'll learn**: API rate limits, throttling behavior, best practices for staying within limits

**Why This Matters:**
- Microsoft Graph has rate limits to protect service performance
- Exceeding limits causes requests to fail with `429 Too Many Requests`
- Understanding limits helps you design robust applications

### Current Rate Limits for Outlook/Mail APIs

**Per User/Mailbox Limits (As of 2025):**
- **10,000 requests per 10 minutes** per app per mailbox
- **4 concurrent requests** maximum per app per mailbox
- **150 MB upload** maximum within 5 minutes per app per mailbox

**Important 2025 Change:**
- Starting **September 30, 2025**: Per-app/per-user per-tenant limit reduced to **half** of total per-tenant limit
- Prevents single user/app from consuming all quota

**Scope:**
- Limits apply per app ID + mailbox combination
- Only Outlook-related APIs count toward this limit
- Non-Outlook resources (OneDrive, etc.) have separate limits

### Current Application Usage Analysis

**Interactive Mode:**
```
Authentication:              1 request
Get user profile:            1 request
Discover shared mailboxes:  47 requests (test access to each)
Discover folders:          308 requests (recursive discovery)
Export emails:        Variable (depends on folder size)
-----------------------------------------------------------
Total baseline:           ~357 requests

Percentage of limit: 3.57% of 10,000 ✅ Well within limits!
```

**Automated Mode (with args):**
```
Authentication:              1 request
Discover folders:          308 requests
Export emails:        Variable
-----------------------------------------------------------
Total baseline:           ~309 requests

Percentage of limit: 3.09% of 10,000 ✅ Safe!
```

**When You Might Hit Limits:**
- Exporting folders with thousands of emails
- Running multiple exports in quick succession
- Batch processing many mailboxes rapidly
- Using pagination to retrieve large datasets

### Throttling Response Behavior

**When throttled, Microsoft Graph returns:**
- HTTP Status Code: `429 Too Many Requests`
- `Retry-After` header: Seconds to wait before retrying
- Error message indicating throttling

**Example Response:**
```
Status: 429 Too Many Requests
Retry-After: 120

{
  "error": {
    "code": "TooManyRequests",
    "message": "The request has been throttled"
  }
}
```

### Best Practices to Avoid Throttling

#### 1. Use Pagination with $top Parameter

**What it does**: Retrieve more items per request, reducing total requests

```csharp
// Bad: Default pagination (10-25 items per page)
var messages = await graphClient.Users[email].Messages.GetAsync();
// May require 100 requests for 1000 messages

// Good: Increase page size
var messages = await graphClient.Users[email].Messages.GetAsync(config =>
{
    config.QueryParameters.Top = 1000;  // Maximum allowed per page
});
// Only 1 request for 1000 messages ✅
```

**Impact**: Can reduce requests by 10-100x for large datasets!

#### 2. Implement Retry Logic with Exponential Backoff

**What it does**: Automatically retry when throttled, respecting Retry-After header

```csharp
private async Task<T> ExecuteWithRetryAsync<T>(Func<Task<T>> operation, int maxRetries = 3)
{
    for (int attempt = 0; attempt <= maxRetries; attempt++)
    {
        try
        {
            return await operation();
        }
        catch (ServiceException ex) when (ex.StatusCode == HttpStatusCode.TooManyRequests)
        {
            if (attempt == maxRetries)
                throw;

            // Get retry delay from response header
            var retryAfter = ex.ResponseHeaders?.RetryAfter?.Delta
                ?? TimeSpan.FromSeconds(Math.Pow(2, attempt)); // Exponential backoff

            Console.WriteLine($"Throttled. Waiting {retryAfter.TotalSeconds} seconds...");
            await Task.Delay(retryAfter);
        }
    }
    throw new Exception("Max retries exceeded");
}

// Usage:
var messages = await ExecuteWithRetryAsync(() =>
    graphClient.Users[email].Messages.GetAsync()
);
```

**Why Exponential Backoff?**
- First retry: Wait 2 seconds
- Second retry: Wait 4 seconds
- Third retry: Wait 8 seconds
- Gives service time to recover

#### 3. Use Batch Requests for Multiple Operations

**What it does**: Combine multiple API calls into single HTTP request

```csharp
// Bad: Multiple individual requests
var inbox = await graphClient.Users[email].MailFolders["Inbox"].GetAsync();
var sent = await graphClient.Users[email].MailFolders["SentItems"].GetAsync();
var drafts = await graphClient.Users[email].MailFolders["Drafts"].GetAsync();
// 3 separate requests

// Good: Single batch request
var batchRequestContent = new BatchRequestContent();

var inboxRequest = graphClient.Users[email].MailFolders["Inbox"].ToGetRequestInformation();
var sentRequest = graphClient.Users[email].MailFolders["SentItems"].ToGetRequestInformation();
var draftsRequest = graphClient.Users[email].MailFolders["Drafts"].ToGetRequestInformation();

var inboxId = batchRequestContent.AddBatchRequestStep(inboxRequest);
var sentId = batchRequestContent.AddBatchRequestStep(sentRequest);
var draftsId = batchRequestContent.AddBatchRequestStep(draftsRequest);

var batchResponse = await graphClient.Batch.PostAsync(batchRequestContent);
// Only 1 request! ✅

// Extract individual responses
var inboxResponse = await batchResponse.GetResponseByIdAsync<MailFolder>(inboxId);
var sentResponse = await batchResponse.GetResponseByIdAsync<MailFolder>(sentId);
var draftsResponse = await batchResponse.GetResponseByIdAsync<MailFolder>(draftsId);
```

**Limitations:**
- Maximum 20 individual requests per batch
- Only 4 requests from batch processed concurrently
- Still counts toward rate limit, but reduces overhead

#### 4. Monitor and Track Request Count

**What it does**: Track your API usage to stay within limits

```csharp
public class RequestTracker
{
    private int _requestCount = 0;
    private DateTime _windowStart = DateTime.UtcNow;
    private readonly int _limitPerWindow = 10000;
    private readonly TimeSpan _window = TimeSpan.FromMinutes(10);

    public async Task<T> TrackRequestAsync<T>(Func<Task<T>> operation)
    {
        // Reset counter if window expired
        if (DateTime.UtcNow - _windowStart > _window)
        {
            _requestCount = 0;
            _windowStart = DateTime.UtcNow;
        }

        // Check if approaching limit (90% threshold)
        if (_requestCount >= _limitPerWindow * 0.9)
        {
            Console.WriteLine($"⚠️  Warning: Approaching rate limit ({_requestCount}/{_limitPerWindow})");

            // Wait until window resets
            var waitTime = _window - (DateTime.UtcNow - _windowStart);
            if (waitTime.TotalSeconds > 0)
            {
                Console.WriteLine($"Waiting {waitTime.TotalSeconds:F0}s for rate limit reset...");
                await Task.Delay(waitTime);
                _requestCount = 0;
                _windowStart = DateTime.UtcNow;
            }
        }

        _requestCount++;
        Console.WriteLine($"Request {_requestCount}/{_limitPerWindow} in current window");

        return await operation();
    }
}

// Usage:
var tracker = new RequestTracker();
var messages = await tracker.TrackRequestAsync(() =>
    graphClient.Users[email].Messages.GetAsync()
);
```

#### 5. Cache Results Locally

**What it does**: Store frequently accessed data to reduce API calls

```csharp
// Cache discovered mailboxes (they don't change often)
public class MailboxCache
{
    private List<MailboxInfo>? _cachedMailboxes;
    private DateTime _cacheExpiry = DateTime.MinValue;
    private readonly TimeSpan _cacheDuration = TimeSpan.FromHours(1);

    public async Task<List<MailboxInfo>> GetMailboxesAsync(Func<Task<List<MailboxInfo>>> fetchFunc)
    {
        // Return cached data if still valid
        if (_cachedMailboxes != null && DateTime.UtcNow < _cacheExpiry)
        {
            Console.WriteLine("Using cached mailboxes ✅");
            return _cachedMailboxes;
        }

        // Fetch fresh data
        Console.WriteLine("Fetching mailboxes from API...");
        _cachedMailboxes = await fetchFunc();
        _cacheExpiry = DateTime.UtcNow.Add(_cacheDuration);

        return _cachedMailboxes;
    }

    public void InvalidateCache()
    {
        _cachedMailboxes = null;
        _cacheExpiry = DateTime.MinValue;
    }
}
```

**What to Cache:**
- ✅ Discovered shared mailboxes (rarely change)
- ✅ Folder structures (relatively static)
- ❌ Email messages (change frequently)
- ❌ Unread counts (change frequently)

#### 6. Implement Progress Throttling

**What it does**: Add delays between requests to spread load

```csharp
// When processing many folders
foreach (var folder in allFolders)
{
    var messages = await graphClient.Users[email]
        .MailFolders[folder.Id]
        .Messages
        .GetAsync();

    // Process messages...

    // Add small delay between folders (100ms)
    await Task.Delay(100);
}

// Or calculate dynamic delay based on rate limit:
var delayMs = (10 * 60 * 1000) / 10000; // 10 minutes / 10,000 requests = 60ms per request
await Task.Delay(delayMs);
```

#### 7. Use $select to Request Only Needed Fields

**What it does**: Reduces response size, improves performance

```csharp
// Bad: Request all fields (larger response, more processing)
var messages = await graphClient.Users[email].Messages.GetAsync();

// Good: Request only needed fields
var messages = await graphClient.Users[email].Messages.GetAsync(config =>
{
    config.QueryParameters.Select = new[]
    {
        "id",
        "subject",
        "from",
        "receivedDateTime"
    };
});
```

**Benefits:**
- Faster response times
- Less bandwidth usage
- Same request count, but more efficient

### Testing Throttling Behavior

**Simulate throttling in development:**

```csharp
// Create test that intentionally hits rate limit
[Test]
public async Task TestThrottlingHandling()
{
    var requestCount = 0;

    try
    {
        // Make rapid requests until throttled
        while (requestCount < 15000) // Exceed 10,000 limit
        {
            await graphClient.Me.GetAsync();
            requestCount++;
        }
    }
    catch (ServiceException ex) when (ex.StatusCode == HttpStatusCode.TooManyRequests)
    {
        Console.WriteLine($"Successfully caught throttling at {requestCount} requests");
        Console.WriteLine($"Retry-After: {ex.ResponseHeaders?.RetryAfter?.Delta}");
        // Test your retry logic here
    }
}
```

### Monitoring in Production

**Log throttling events:**

```csharp
try
{
    await graphClient.Users[email].Messages.GetAsync();
}
catch (ServiceException ex) when (ex.StatusCode == HttpStatusCode.TooManyRequests)
{
    // Log to file, monitoring service, etc.
    var log = new
    {
        Timestamp = DateTime.UtcNow,
        Operation = "GetMessages",
        Mailbox = email,
        RetryAfter = ex.ResponseHeaders?.RetryAfter?.Delta,
        ErrorMessage = ex.Message
    };

    File.AppendAllText("throttling.log", JsonSerializer.Serialize(log) + "\n");

    // Then retry...
}
```

### Summary: Rate Limit Checklist

When building features that make many API calls:

- ✅ Use `$top=1000` for pagination
- ✅ Implement retry logic with `Retry-After` header
- ✅ Consider batch requests for multiple operations
- ✅ Track request count in your app
- ✅ Cache data when appropriate
- ✅ Add delays between requests if processing many items
- ✅ Use `$select` to request only needed fields
- ✅ Log throttling events for monitoring
- ✅ Test throttling behavior in development

**Your current app is well within limits, but these practices become important as you add features like large-scale exports or batch processing!**

---

## Testing & Debugging

### Chrome DevTools Equivalent: Visual Studio / VS Code Debugger

**Setting Breakpoints**:
1. Click left margin in code editor (red dot appears)
2. Run with F5 (Debug mode)
3. Execution pauses at breakpoint
4. Inspect variables, step through code

**Debugging Async Code**:
- Set breakpoint on `await` line
- Step over (F10) to wait for completion
- Step into (F11) to debug async method

### Console Debugging

**Add Diagnostic Output**:
```csharp
Console.WriteLine($"Debug: ClientId = {clientId}");
Console.WriteLine($"Debug: TenantId = {tenantId}");
Console.WriteLine($"Debug: Folders count = {folders?.Value?.Count}");
```

**Try-Catch for Detailed Errors**:
```csharp
try
{
    var user = await graphClient.Me.GetAsync();
}
catch (Exception ex)
{
    Console.WriteLine($"Error: {ex.Message}");
    if (ex.InnerException != null)
    {
        Console.WriteLine($"Inner Error: {ex.InnerException.Message}");
    }
    Console.WriteLine($"Stack Trace: {ex.StackTrace}");
}
```

### Azure Portal Monitoring

**Check Authentication Logs**:
1. Azure Portal → Azure Active Directory
2. "Sign-in logs" or "Enterprise applications"
3. Find your app
4. View authentication attempts
5. See what permissions were used

**Check App Registration**:
1. Verify permissions are granted
2. Check admin consent status
3. Review authentication settings

### Network Debugging with Fiddler

**Optional**: Use Fiddler to see HTTP requests:
1. Install Fiddler
2. Run app
3. View all Graph API requests/responses
4. Useful for understanding API calls

---

## Troubleshooting Guide

### Error 1: "The mailbox is either inactive, soft-deleted, or is hosted on-premise"

**Symptom**:
- Authentication succeeds
- Getting user profile works
- Accessing mailbox fails

**Root Cause**: Tenant ID mismatch

**Solution**:

**For Personal Accounts**:
```json
{
  "AzureAd": {
    "ClientId": "your-client-id",
    "TenantId": "consumers"  ← Must be "consumers"
  }
}
```

**For Organizational Accounts**:
```json
{
  "AzureAd": {
    "ClientId": "your-client-id",
    "TenantId": "0b474a1c-..."  ← Your org's tenant ID
  }
}
```

**How to Fix**:
1. Identify your account type
2. Update `appsettings.Development.json`
3. Rebuild and run

### Error 2: "Needs administrator approval" / "Necessita de aprovação do administrador"

**Symptom**:
- Organizational account
- Authentication starts
- Error during consent

**Root Cause**: Organizational tenant requires admin consent

**Solution**:
1. Provide `ADMIN_SETUP_GUIDE.md` to IT admin
2. Admin grants consent in Azure Portal:
   - API permissions → Grant admin consent
3. Wait 5-10 minutes for propagation
4. Try again

**Alternative (if you're admin)**:
1. Azure Portal → Your app registration
2. API permissions
3. "Grant admin consent for [Organization]"
4. Confirm

### Error 3: "Insufficient privileges to complete the operation"

**Symptom**:
- Accessing shared mailbox fails
- Own mailbox works fine

**Root Cause**: No permission to access that mailbox

**Solution**:
- Verify you have "Full Access" permission
- Contact Exchange administrator
- Request delegation for that mailbox

### Error 4: Configuration file not found

**Symptom**: "Could not load file or assembly..."

**Root Cause**: Missing appsettings file

**Solution**:
1. Copy `appsettings.Example.json` to `appsettings.Development.json`
2. Fill in actual values
3. Verify file is in same directory as Program.cs

### Build Errors

**Missing NuGet Packages**:
```bash
dotnet restore
dotnet build
```

**Version Conflicts**:
```bash
dotnet clean
dotnet restore --force
dotnet build
```

### Runtime Errors

**"DeviceCodeCredential authentication failed"**:
- Check internet connection
- Verify Client ID is correct
- Check Tenant ID is correct
- Ensure public client flows enabled

**"Invalid JSON"** in output file:
- Check JsonSerializerOptions
- Verify all properties are serializable
- Look for circular references

---

## Additional Resources

### Official Documentation

- [Microsoft Graph API](https://docs.microsoft.com/en-us/graph/)
- [Graph SDK for .NET](https://docs.microsoft.com/en-us/graph/sdks/sdks-overview)
- [Azure Identity Library](https://docs.microsoft.com/en-us/dotnet/api/overview/azure/identity-readme)
- [Device Code Flow](https://docs.microsoft.com/en-us/azure/active-directory/develop/v2-oauth2-device-code)
- [Microsoft Graph Permissions](https://docs.microsoft.com/en-us/graph/permissions-reference)

### Tools

- [Graph Explorer](https://developer.microsoft.com/en-us/graph/graph-explorer) - Test Graph API calls
- [Azure Portal](https://portal.azure.com) - Manage app registrations
- [JWT.ms](https://jwt.ms) - Decode access tokens
- [Fiddler](https://www.telerik.com/fiddler) - HTTP debugging

### Learning Resources

- [Microsoft Graph Tutorials](https://docs.microsoft.com/en-us/graph/tutorials)
- [Azure AD App Registration Guide](https://docs.microsoft.com/en-us/azure/active-directory/develop/quickstart-register-app)
- [OAuth 2.0 Explained](https://oauth.net/2/)

### Best Practices

1. **Never hardcode credentials** - Use configuration files
2. **Use .gitignore** - Prevent credential leaks
3. **Least privilege** - Request only necessary permissions
4. **Error handling** - Always use try-catch with async calls
5. **Logging** - Add console output for debugging
6. **Testing** - Test after each major change
7. **Documentation** - Keep README and PROJECT_SUMMARY updated

---

## Account Types Comparison

| Feature | Personal Account | Organizational Account |
|---------|-----------------|------------------------|
| **Tenant ID** | `"consumers"` | Specific tenant ID |
| **Admin Consent** | ❌ Not required | ✅ Required |
| **Setup Complexity** | Simple | Moderate (needs admin) |
| **Authentication** | Immediate | Requires admin approval first |
| **Shared Mailboxes** | Not applicable | Available (with permissions) |
| **Use Case** | Personal mailbox export | Enterprise/work mailbox export |
| **IT Involvement** | None | Administrator needed |

---

## What's Next?

### Completed Enhancements (October 2025)

1. ✅ **Configurable Email Count** (October 20, 2025)
   - Command-line argument: `-c` / `--count`
   - Three modes:
     - Default: 5 emails (no argument)
     - Specific count: `-c 100` exports 100 emails
     - All emails: `-c 0` exports entire folder using pagination
   - Implemented PageIterator for large datasets
   - Progress indicators for large exports

2. ✅ **Configuration-based Known Mailboxes** (October 20, 2025)
   - Removed hardcoded mailbox from code
   - Added `KnownMailboxes` section to appsettings.json
   - Supports multiple known mailboxes
   - Better separation of code and configuration

3. ✅ **Folder Pagination Fix** (October 20, 2025)
   - **Critical Bug Fix**: Was only retrieving first page of folders
   - Added `Top = 999` parameter to folder requests
   - Implemented PageIterator for root folders and child folders
   - Result: 308 → 1,445 folders discovered (4.7x increase!)
   - Ensures ALL folders are discovered, not just first page

4. ✅ **Folder Discovery Early Exit Optimization** (October 21, 2025)
   - Stop searching once target folder is found
   - Added `folderFound` flag to halt enumeration
   - Skip printing all folders when searching for specific folder
   - Performance improvement: 932 folders vs 1,445 (35% faster)
   - Significantly speeds up automated/scripted exports

### Future Enhancements

1. **Interactive Folder Selection**
   - Menu system for folder selection
   - Export multiple folders in one run

2. **Date Range Filtering**
   - Filter by date received
   - Export only recent emails
   - Archive old emails separately

3. **Attachment Support**
   - Download attachments
   - Save to separate folder
   - Include attachment metadata

4. **Multiple Output Formats**
   - CSV export
   - XML export
   - Individual .eml files
   - Excel workbook

5. **Scheduled Exports**
   - Run on schedule (Windows Task Scheduler)
   - Incremental exports (only new emails)
   - Email notifications on completion

6. **Search/Filter**
   - Filter by sender
   - Filter by subject keywords
   - Filter by importance/category

7. **Retry Logic and Throttling Protection**
   - Implement exponential backoff for rate limit errors
   - Request count tracking
   - Proactive throttling before hitting limits

### Experimental Ideas for Future Exploration

1. **Enhanced Console UI with Spectre.Console**
   - **Library**: [Spectre.Console](https://github.com/spectreconsole/spectre.console)
   - **What it provides**:
     - Rich text formatting with colors and styles
     - Interactive menus and prompts
     - Progress bars and spinners
     - Tables with borders and formatting
     - Tree views for hierarchical data (perfect for folders!)
     - Panels and layouts
   - **Why explore this**:
     - Current app uses basic `Console.WriteLine`
     - Spectre.Console would make output more professional
     - Better user experience with interactive menus
     - Visual progress indicators for long operations
     - Folder selection could use tree view instead of numbered list
   - **Learning goals**:
     - Modern console UI development
     - User experience design in CLI apps
     - Understanding markup languages for console output
   - **Example use cases**:
     ```csharp
     // Instead of basic menu:
     Console.WriteLine("[1] Inbox");
     Console.WriteLine("[2] Sent Items");

     // Use Spectre.Console:
     var folder = AnsiConsole.Prompt(
         new SelectionPrompt<string>()
             .Title("Select [green]folder[/] to export:")
             .AddChoices(new[] { "Inbox", "Sent Items", "Drafts" })
     );

     // Instead of basic progress:
     Console.WriteLine("Exporting...");

     // Use Spectre.Console:
     AnsiConsole.Progress()
         .Start(ctx => {
             var task = ctx.AddTask("[green]Exporting emails[/]");
             // Update progress as emails are exported
         });

     // Display folders as tree:
     var tree = new Tree("Mailbox Folders");
     tree.AddNode("[blue]Inbox[/]")
         .AddNode("[blue]Clients[/]")
             .AddNode("[blue]A[/]")
                 .AddNode("[blue]Aber[/]");
     AnsiConsole.Write(tree);
     ```

2. **Code Modularization & Architecture Refactoring**
   - **Current state**: ~400+ lines in single `Program.cs` file
   - **Problem**: As features grow, single file becomes hard to maintain
   - **Goals**:
     - Separate concerns (auth, API calls, export, UI)
     - Improve testability
     - Enable code reuse
     - Make adding features easier
   - **Proposed architecture**:
     ```
     OutlookExporter/
     ├── Program.cs                    # Entry point (~50 lines)
     ├── Services/
     │   ├── AuthenticationService.cs  # Device Code Flow, token management
     │   ├── GraphApiService.cs        # All Graph API calls
     │   ├── MailboxDiscovery.cs       # Shared mailbox discovery
     │   ├── FolderDiscovery.cs        # Recursive folder discovery
     │   └── EmailExporter.cs          # Export logic, JSON serialization
     ├── Models/
     │   ├── MailboxInfo.cs            # (DisplayName, Email, Type)
     │   ├── FolderInfo.cs             # (Id, Name, Path, Total, Unread)
     │   └── ExportOptions.cs          # Command-line args, config
     ├── UI/
     │   ├── ConsoleUI.cs              # User interaction, menus
     │   └── ArgumentParser.cs         # Command-line parsing
     └── Configuration/
         └── AppSettings.cs            # Configuration model
     ```
   - **Benefits**:
     - ✅ Single Responsibility Principle (each class has one job)
     - ✅ Dependency Injection (easier testing, mocking)
     - ✅ Unit testable (can test each service independently)
     - ✅ Easier to add features (know where to add code)
     - ✅ Better for team collaboration
     - ✅ Reusable components (use GraphApiService in other projects)
   - **Learning goals**:
     - Software architecture patterns
     - SOLID principles
     - Dependency injection in .NET
     - Unit testing with xUnit/NUnit
     - Separation of concerns
   - **Example refactoring**:
     ```csharp
     // Before (in Program.cs):
     var users = await graphClient.Users.GetAsync(...);
     foreach (var user in users.Value) { ... }

     // After (in MailboxDiscoveryService.cs):
     public class MailboxDiscoveryService
     {
         private readonly IGraphApiService _graphApi;

         public MailboxDiscoveryService(IGraphApiService graphApi)
         {
             _graphApi = graphApi;
         }

         public async Task<List<MailboxInfo>> DiscoverSharedMailboxesAsync()
         {
             var users = await _graphApi.GetDisabledUsersAsync();
             var accessible = new List<MailboxInfo>();

             foreach (var user in users)
             {
                 if (await _graphApi.TestMailboxAccessAsync(user.Mail))
                 {
                     accessible.Add(new MailboxInfo(
                         user.DisplayName,
                         user.Mail,
                         MailboxType.Shared
                     ));
                 }
             }

             return accessible;
         }
     }

     // Usage in Program.cs:
     var discoveryService = new MailboxDiscoveryService(graphApiService);
     var mailboxes = await discoveryService.DiscoverSharedMailboxesAsync();
     ```
   - **Testing example**:
     ```csharp
     [Fact]
     public async Task DiscoverSharedMailboxes_ReturnsOnlyAccessible()
     {
         // Arrange
         var mockGraphApi = new Mock<IGraphApiService>();
         mockGraphApi.Setup(x => x.GetDisabledUsersAsync())
             .ReturnsAsync(new[] { user1, user2, user3 });
         mockGraphApi.Setup(x => x.TestMailboxAccessAsync("user1@test.com"))
             .ReturnsAsync(true);
         mockGraphApi.Setup(x => x.TestMailboxAccessAsync("user2@test.com"))
             .ReturnsAsync(false);

         var service = new MailboxDiscoveryService(mockGraphApi.Object);

         // Act
         var result = await service.DiscoverSharedMailboxesAsync();

         // Assert
         Assert.Single(result);
         Assert.Equal("user1@test.com", result[0].Email);
     }
     ```

### Other Microsoft 365 Integrations

Now that you understand Microsoft Graph:
- Export calendar events
- Access OneDrive files
- Read Teams messages
- Access SharePoint lists
- Manage contacts

---

## Success Criteria

By completing this learning plan, you should be able to:

✅ Register applications in Azure AD
✅ Configure API permissions and understand delegation
✅ Implement OAuth 2.0 Device Code Flow
✅ Call Microsoft Graph API endpoints
✅ Handle different account types (personal vs organizational)
✅ Troubleshoot common authentication errors
✅ Understand tenant IDs and admin consent
✅ Write async/await C# code
✅ Serialize data to JSON
✅ Implement secure configuration practices
✅ Access shared mailboxes
✅ Build production-ready console applications

---

**Congratulations on completing the Outlook Email Exporter! You now have the foundation to build any Microsoft 365 integration using Microsoft Graph API.**

**Remember: The best way to solidify your learning is to build something new. Try extending this project or creating a different Microsoft 365 tool!**

---

*Document created: October 17, 2025*
*Project duration: 4-5 hours*
*Lines of code: ~240 lines*
*Status: Production-ready*
