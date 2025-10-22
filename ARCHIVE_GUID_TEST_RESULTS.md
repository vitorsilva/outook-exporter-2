# Archive GUID Test Results

**Date:** 2025-10-22
**Test Subject:** Can Microsoft Graph API access archive mailboxes using ArchiveGuid?

## Test Configuration

- **Archive GUID Tested:** `d88b8107-7f31-4e0a-999b-723a8ba54ac0`
- **Primary Mailbox:** `Grupo.ComDev@samsys.pt`
- **Display Name:** Arquivo ComDev
- **Graph API Version:** v1.0
- **Authentication:** Device Code Flow with delegated permissions

## Test Results

### Test 1: Using ArchiveGuid as User Identifier ❌

**Attempted:**
```
GET https://graph.microsoft.com/v1.0/users/{archiveGuid}/mailFolders
GET https://graph.microsoft.com/v1.0/users/{archiveGuid}
GET https://graph.microsoft.com/v1.0/users/{archiveGuid}/messages
```

**Result:** ❌ **FAILED**

**Error Messages:**
- `The requested user 'd88b8107-7f31-4e0a-999b-723a8ba54ac0' is invalid.`
- `Resource 'd88b8107-7f31-4e0a-999b-723a8ba54ac0' does not exist`

**Conclusion:** ArchiveGuid is **NOT** recognized as a valid user identifier in Microsoft Graph API. The Graph API only accepts:
- Azure AD Object ID (user's ID in Azure AD, not Exchange GUID)
- UserPrincipalName (UPN)
- Email address (mail property)

### Test 2: Using ArchiveMsgFolderRoot ❌ (But Revealing!)

**Attempted:**
```
GET https://graph.microsoft.com/v1.0/users/Grupo.ComDev@samsys.pt/mailFolders/ArchiveMsgFolderRoot
```

**Result:** ❌ **FAILED** (But with a very interesting error!)

**Error Message:**
```
Item 'ArchiveMsgFolderRoot' doesn't belong to the targeted mailbox '50bc83e9-5d9e-462b-950c-4eec6892abfa'.
The item exists in an archive mailbox.
```

**This is FASCINATING because:**
1. ✅ Graph API **KNOWS** the archive mailbox exists
2. ✅ Graph API **FOUND** the ArchiveMsgFolderRoot
3. ✅ Graph API **RECOGNIZED** it's in an archive mailbox
4. ❌ But it **REFUSES** to return it (deliberate block, not a "not found" error)

**Mailbox GUID Revealed:** The error message exposed the primary mailbox's internal GUID: `50bc83e9-5d9e-462b-950c-4eec6892abfa`

## Key Findings

### 1. ArchiveGuid is Exchange-Specific

The `ArchiveGuid` property from PowerShell's `Get-Mailbox` is an **Exchange Online property**, not an Azure AD property. Graph API uses Azure AD identifiers, which is why it doesn't recognize the ArchiveGuid.

### 2. Graph API Can "See" Archives but Blocks Access

The error message "The item exists in an archive mailbox" proves that:
- Graph API is **aware** of archive mailboxes
- The blocking is **intentional**, not due to missing functionality
- This is a **policy/permission restriction**, not a technical limitation

### 3. ArchiveMsgFolderRoot is "Known" to Graph API

Even though it's not in the official well-known folder list, Graph API:
- Recognizes `ArchiveMsgFolderRoot` as a valid folder identifier
- Can locate it in the mailbox structure
- But enforces a restriction that prevents cross-mailbox access

## Implications

### What This Means:

1. **ArchiveGuid Cannot Be Used**
   - Don't use ArchiveGuid as a user identifier in Graph API
   - It's not equivalent to Azure AD Object ID

2. **Archives Are Technically Accessible (Maybe)**
   - The error suggests Graph API **can** access archives internally
   - The restriction might be permission-based or API-version-based
   - There might be an undocumented way to access them with special permissions

3. **Microsoft's Intentional Restriction**
   - This is a deliberate design decision, not a technical gap
   - Microsoft may add official archive support in the future
   - Or they may continue blocking it to push users toward EWS/PowerShell

## Recommendations

### For Archive Access:

1. **Use Exchange Online PowerShell** (Most Reliable)
   ```powershell
   Connect-ExchangeOnline
   $mailbox = Get-Mailbox -Identity "Grupo.ComDev@samsys.pt"
   $archiveStats = Get-MailboxStatistics -Identity $mailbox.ArchiveGuid -Archive
   ```

2. **Use EWS (Exchange Web Services)** (Programmatic Access)
   - Can access `ArchiveMsgFolderRoot` via EWS
   - Being deprecated but still works
   - Use EWS Managed API 2.2 with OAuth

3. **Hybrid Approach** (Recommended for This Project)
   - Use PowerShell to export archive information (including ArchiveGuid)
   - Store archive email addresses in `appsettings.json` with `IsArchive: true`
   - Access archives by email address (if naming pattern is known)

### For This Application:

Continue using the **heuristic naming pattern approach** with configuration-based known archives. The test confirms there's no "magic bullet" using ArchiveGuid in Graph API.

## Test Command

To reproduce these results:
```bash
dotnet run --project OutlookExporter -- --test-archive d88b8107-7f31-4e0a-999b-723a8ba54ac0 -m Grupo.ComDev@samsys.pt
```

## References

- [Microsoft Graph Mail API Overview](https://learn.microsoft.com/en-us/graph/api/resources/mail-api-overview?view=graph-rest-1.0)
- [Accessing Archive Mailboxes Blog Post](https://devblogs.microsoft.com/microsoft365dev/accessing-outlook-items-in-a-users-archived-shared-or-delegated-mailboxes-using-microsoft-graph/)
- PowerShell `Get-Mailbox` cmdlet documentation

## Conclusion

**ArchiveGuid cannot be used to access archive mailboxes via Microsoft Graph API.**

However, the error messages revealed that Graph API is **aware** of archive mailboxes and deliberately restricts access to them. This confirms Microsoft's documented limitation that "The API does not support accessing in-place archive mailboxes."

The most reliable approach remains using **Exchange Online PowerShell** for archive detection and access.
