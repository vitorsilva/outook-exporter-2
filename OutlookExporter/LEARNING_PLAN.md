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

## Phase 5: Multi-Mailbox Support (1 hour)

### Goal
Access shared/delegated mailboxes beyond primary mailbox.

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

### Step 5.2: Implement Mailbox Discovery

**What you'll learn**: User input, menu system, dynamic endpoint selection

**Code to Add**:
```csharp
Console.WriteLine("\n" + new string('=', 50));
Console.WriteLine("Discovering available mailboxes...");
Console.WriteLine(new string('=', 50));

var availableMailboxes = new List<(string DisplayName, string Email, string Type)>();

// Add primary mailbox
availableMailboxes.Add((
    user?.DisplayName ?? "Primary Mailbox",
    user?.Mail ?? user?.UserPrincipalName ?? "",
    "Primary"
));

Console.WriteLine($"\nFound {availableMailboxes.Count} mailbox(es):");
for (int i = 0; i < availableMailboxes.Count; i++)
{
    Console.WriteLine($"  [{i + 1}] {availableMailboxes[i].DisplayName} ({availableMailboxes[i].Email}) - {availableMailboxes[i].Type}");
}
Console.WriteLine($"  [0] Enter custom mailbox email address");
```

### Step 5.3: Implement Mailbox Selection

**Code to Add**:
```csharp
Console.Write("\nSelect mailbox (enter number): ");
var selection = Console.ReadLine();

string selectedMailboxEmail = "";
string selectedMailboxName = "";

if (int.TryParse(selection, out int selectedIndex))
{
    if (selectedIndex == 0)
    {
        Console.Write("Enter mailbox email address: ");
        selectedMailboxEmail = Console.ReadLine() ?? "";
        selectedMailboxName = selectedMailboxEmail;
    }
    else if (selectedIndex > 0 && selectedIndex <= availableMailboxes.Count)
    {
        selectedMailboxEmail = availableMailboxes[selectedIndex - 1].Email;
        selectedMailboxName = availableMailboxes[selectedIndex - 1].DisplayName;
    }
    else
    {
        Console.WriteLine("Invalid selection, using primary mailbox.");
        selectedMailboxEmail = user?.Mail ?? user?.UserPrincipalName ?? "";
        selectedMailboxName = user?.DisplayName ?? "Primary";
    }
}

Console.WriteLine($"\nSelected mailbox: {selectedMailboxName} ({selectedMailboxEmail})");
```

### Step 5.4: Update API Calls to Use Selected Mailbox

**Code Changes**:

**Before**:
```csharp
var folders = await graphClient.Me.MailFolders.GetAsync();
var messages = await graphClient.Me.MailFolders["Inbox"].Messages.GetAsync();
```

**After**:
```csharp
var folders = await graphClient.Users[selectedMailboxEmail].MailFolders.GetAsync();
var messages = await graphClient.Users[selectedMailboxEmail].MailFolders["Inbox"].Messages.GetAsync();
```

**Why This Works**:
- `.Users[email]` can be used for any mailbox user has access to
- Includes own mailbox (user's email)
- Includes shared mailboxes (with permissions)
- Same API structure for all mailboxes

**Testing**:
1. Test with primary mailbox (should work as before)
2. Test with shared mailbox email (if available)
3. Test with invalid email (should get permission error)

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

### Future Enhancements

1. **Interactive Folder Selection**
   - Let user choose which folder to export
   - Menu system for folder selection
   - Export multiple folders

2. **Date Range Filtering**
   - Filter by date received
   - Export only recent emails
   - Archive old emails separately

3. **Pagination**
   - Handle large mailboxes (thousands of emails)
   - Process in batches
   - Progress indicators

4. **Attachment Support**
   - Download attachments
   - Save to separate folder
   - Include attachment metadata

5. **Multiple Output Formats**
   - CSV export
   - XML export
   - Individual .eml files
   - Excel workbook

6. **Scheduled Exports**
   - Run on schedule (Windows Task Scheduler)
   - Incremental exports (only new emails)
   - Email notifications on completion

7. **Search/Filter**
   - Filter by sender
   - Filter by subject keywords
   - Filter by importance/category

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
