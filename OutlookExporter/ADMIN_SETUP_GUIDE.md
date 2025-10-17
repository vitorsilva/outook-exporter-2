# Admin Setup Guide - Outlook Email Exporter

**For System Administrators**

This document provides instructions for configuring the Outlook Email Exporter application in your organization's Azure Active Directory tenant.

## Application Information

**Application Name:** Outlook Email Exporter

**Application (Client) ID:** `5723b5d0-bf95-4e8f-97f4-dbaf30a9fad9`

**Tenant ID:** `0b474a1c-e4d1-477f-95cb-9a74ddada3a3`

**Purpose:** Read-only email export tool for Microsoft 365 mailboxes

## Required Configuration Steps

### 1. Locate the App Registration

1. Sign in to [Azure Portal](https://portal.azure.com) with admin credentials
2. Navigate to: **Azure Active Directory** → **App registrations**
3. Find the app: "Outlook Email Exporter"
   - Or search by Client ID: `5723b5d0-bf95-4e8f-97f4-dbaf30a9fad9`

### 2. Configure API Permissions

1. In the app registration, click **API permissions** in the left menu
2. Verify the following **Delegated permissions** are present:

   | Permission | Purpose |
   |------------|---------|
   | `User.Read` | Read user profile information |
   | `Mail.Read` | Read user's mailbox (emails) |
   | `Mail.ReadBasic` | Basic mail access |
   | `MailboxSettings.Read` | Read mailbox settings |

3. If any permissions are missing:
   - Click **+ Add a permission**
   - Select **Microsoft Graph**
   - Select **Delegated permissions**
   - Search for and add the missing permissions

### 3. Grant Admin Consent

**This is the critical step that allows users to use the application:**

1. In the **API permissions** page
2. Click the button: **"Grant admin consent for [Your Organization Name]"**
3. Confirm the consent dialog
4. All permissions should now show a green checkmark under "Status"

### 4. Enable Public Client Flow

1. In the app registration, click **Authentication** in the left menu
2. Scroll down to **Advanced settings**
3. Under **"Allow public client flows"**, set toggle to **Yes**
4. Click **Save** at the top of the page

## Security & Compliance Information

### Application Security Profile

✅ **Read-Only Access:** Cannot send, modify, or delete emails
✅ **User Consent:** Only accesses data when user is signed in
✅ **No Application Permissions:** Cannot access data without user interaction
✅ **Standard Microsoft API:** Uses official Microsoft Graph API
✅ **Device Code Flow:** Secure OAuth 2.0 authentication method

### Data Privacy

- Application does **NOT** store credentials
- Authentication tokens are managed by Microsoft Identity Platform
- Email data is exported to local files on the user's machine
- No data is sent to external servers
- Application runs locally on user's workstation

### Compliance Considerations

- **GDPR Compliant:** User controls their own data export
- **No Cloud Storage:** All exports remain on user's local machine
- **Audit Trail:** Azure AD logs all authentication and consent events
- **Revocable:** Users can revoke app access at any time via their Microsoft account settings

## Verification Steps

After completing the configuration:

1. User should be able to authenticate without "admin approval required" error
2. Application will request user consent on first use
3. User can view/revoke consent in their Microsoft account settings
4. Authentication events will appear in Azure AD sign-in logs

## Common Questions

### Q: Why does this app need admin consent?

**A:** Your organization's Azure AD is configured to require admin approval for applications requesting delegated permissions. This is a security best practice for organizational tenants.

### Q: Can users revoke access?

**A:** Yes, users can revoke the app's access at any time through their Microsoft account settings or via Azure AD My Apps portal.

### Q: Will this work for all users in the organization?

**A:** Yes, once admin consent is granted, all users in the tenant can authenticate and use the application.

### Q: What about guest users?

**A:** Guest users (external accounts) in your tenant will need to use their own tenant's app registration. They cannot access mailboxes through your tenant's consent.

### Q: Is there any cost?

**A:** No, this uses standard Microsoft Graph API which is included in Microsoft 365 licenses. No additional costs.

## Troubleshooting

### Users still see "needs admin approval"

- Verify admin consent was successfully granted (green checkmarks on permissions)
- Wait 5-10 minutes for consent to propagate
- Ask user to clear browser cache and try again
- Check that user is signing in with correct organizational account

### "Public client flow not allowed" error

- Ensure "Allow public client flows" is set to **Yes** in Authentication settings
- Click Save after changing the setting

### Permission errors after consent

- Verify all 4 required permissions are present
- Ensure permissions are **Delegated** (not Application)
- Re-grant admin consent if permissions were modified

## Support Information

**Application Developer:** [Your Organization/Department]

**Technical Documentation:** See `README.md` and `PROJECT_SUMMARY.md` in project files

**For Questions:** Contact the user who requested this setup

---

## Quick Reference: Required Permissions

```
Microsoft Graph - Delegated Permissions:
  ✓ User.Read
  ✓ Mail.Read
  ✓ Mail.ReadBasic
  ✓ MailboxSettings.Read

Advanced Settings:
  ✓ Allow public client flows: Yes

Admin Action Required:
  ✓ Grant admin consent for organization
```

---

**Last Updated:** October 17, 2025
