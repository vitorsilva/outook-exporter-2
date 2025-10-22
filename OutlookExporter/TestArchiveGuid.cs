using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;

namespace OutlookExporter;

/// <summary>
/// Test utility to verify if ArchiveGuid can be used to access archive mailboxes via Graph API.
/// This is an experimental test - ArchiveGuid is not documented as a valid user identifier.
/// </summary>
public class ArchiveGuidTester
{
    public static async Task<bool> TestArchiveGuidAccess(
        GraphServiceClient graphClient,
        string archiveGuid,
        string displayName,
        string primarySmtpAddress)
    {
        Console.WriteLine("\n" + new string('=', 70));
        Console.WriteLine("TESTING ARCHIVE GUID ACCESS (Experimental)");
        Console.WriteLine(new string('=', 70));
        Console.WriteLine($"Display Name: {displayName}");
        Console.WriteLine($"Primary SMTP: {primarySmtpAddress}");
        Console.WriteLine($"Archive GUID: {archiveGuid}");
        Console.WriteLine(new string('=', 70));

        bool successfullyAccessed = false;

        // Test 1: Try to access mailFolders using ArchiveGuid
        Console.WriteLine("\n[Test 1] Attempting: /users/{archiveGuid}/mailFolders");
        try
        {
            var folders = await graphClient.Users[archiveGuid].MailFolders.GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Top = 5;
            });

            if (folders?.Value != null && folders.Value.Count > 0)
            {
                Console.WriteLine($"✓ SUCCESS! Found {folders.Value.Count} folder(s):");
                foreach (var folder in folders.Value.Take(3))
                {
                    Console.WriteLine($"  - {folder.DisplayName} ({folder.TotalItemCount} items)");
                }
                if (folders.Value.Count > 3)
                {
                    Console.WriteLine($"  ... and {folders.Value.Count - 3} more folders");
                }
                successfullyAccessed = true;
            }
            else
            {
                Console.WriteLine("✗ Request succeeded but returned no folders");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ FAILED: {ex.Message}");
            if (ex.Message.Contains("404") || ex.Message.Contains("Not Found"))
            {
                Console.WriteLine("  → ArchiveGuid is NOT a valid user identifier");
            }
            else if (ex.Message.Contains("403") || ex.Message.Contains("Forbidden"))
            {
                Console.WriteLine("  → Permission denied (may need additional scopes)");
            }
        }

        // Test 2: Try to get user info using ArchiveGuid
        Console.WriteLine("\n[Test 2] Attempting: /users/{archiveGuid} (get user info)");
        try
        {
            var user = await graphClient.Users[archiveGuid].GetAsync();

            if (user != null)
            {
                Console.WriteLine($"✓ SUCCESS! Retrieved user info:");
                Console.WriteLine($"  - DisplayName: {user.DisplayName}");
                Console.WriteLine($"  - Mail: {user.Mail}");
                Console.WriteLine($"  - UserPrincipalName: {user.UserPrincipalName}");
                successfullyAccessed = true;
            }
            else
            {
                Console.WriteLine("✗ Request succeeded but returned no user info");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ FAILED: {ex.Message}");
        }

        // Test 3: Try to get messages using ArchiveGuid
        Console.WriteLine("\n[Test 3] Attempting: /users/{archiveGuid}/messages");
        try
        {
            var messages = await graphClient.Users[archiveGuid].Messages.GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Top = 5;
            });

            if (messages?.Value != null && messages.Value.Count > 0)
            {
                Console.WriteLine($"✓ SUCCESS! Found {messages.Value.Count} message(s):");
                foreach (var msg in messages.Value.Take(3))
                {
                    Console.WriteLine($"  - {msg.Subject} (from: {msg.From?.EmailAddress?.Address})");
                }
                successfullyAccessed = true;
            }
            else
            {
                Console.WriteLine("✗ Request succeeded but returned no messages");
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ FAILED: {ex.Message}");
        }

        // Summary
        Console.WriteLine("\n" + new string('=', 70));
        if (successfullyAccessed)
        {
            Console.WriteLine("✓✓✓ BREAKTHROUGH! ArchiveGuid CAN be used to access mailbox data!");
            Console.WriteLine("This is an undocumented feature that could revolutionize archive access.");
        }
        else
        {
            Console.WriteLine("✗✗✗ ArchiveGuid cannot be used as a user identifier in Graph API");
            Console.WriteLine("This confirms Microsoft's limitation: archives aren't accessible via Graph API");
        }
        Console.WriteLine(new string('=', 70));

        return successfullyAccessed;
    }

    /// <summary>
    /// Test if ArchiveMsgFolderRoot works for accessing archive content
    /// </summary>
    public static async Task<bool> TestArchiveMsgFolderRoot(
        GraphServiceClient graphClient,
        string primarySmtpAddress)
    {
        Console.WriteLine("\n" + new string('=', 70));
        Console.WriteLine("TESTING ArchiveMsgFolderRoot APPROACH");
        Console.WriteLine(new string('=', 70));
        Console.WriteLine($"Primary SMTP: {primarySmtpAddress}");
        Console.WriteLine(new string('=', 70));

        bool successfullyAccessed = false;

        Console.WriteLine("\nAttempting: /users/{primarySmtp}/mailFolders/ArchiveMsgFolderRoot");
        try
        {
            var archiveRoot = await graphClient.Users[primarySmtpAddress]
                .MailFolders["ArchiveMsgFolderRoot"]
                .GetAsync();

            if (archiveRoot != null)
            {
                Console.WriteLine($"✓ SUCCESS! Archive root folder found:");
                Console.WriteLine($"  - DisplayName: {archiveRoot.DisplayName}");
                Console.WriteLine($"  - TotalItemCount: {archiveRoot.TotalItemCount}");
                Console.WriteLine($"  - ChildFolderCount: {archiveRoot.ChildFolderCount}");
                successfullyAccessed = true;

                // Try to get child folders
                Console.WriteLine("\nAttempting to get child folders...");
                try
                {
                    var childFolders = await graphClient.Users[primarySmtpAddress]
                        .MailFolders["ArchiveMsgFolderRoot"]
                        .ChildFolders
                        .GetAsync();

                    if (childFolders?.Value != null && childFolders.Value.Count > 0)
                    {
                        Console.WriteLine($"✓ Found {childFolders.Value.Count} child folder(s):");
                        foreach (var folder in childFolders.Value.Take(5))
                        {
                            Console.WriteLine($"  - {folder.DisplayName} ({folder.TotalItemCount} items)");
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"✗ Failed to get child folders: {ex.Message}");
                }
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ FAILED: {ex.Message}");
            if (ex.Message.Contains("404") || ex.Message.Contains("ErrorItemNotFound"))
            {
                Console.WriteLine("  → ArchiveMsgFolderRoot is not a recognized well-known folder");
            }
        }

        Console.WriteLine("\n" + new string('=', 70));
        if (successfullyAccessed)
        {
            Console.WriteLine("✓✓✓ ArchiveMsgFolderRoot works! Archives can be accessed this way!");
        }
        else
        {
            Console.WriteLine("✗✗✗ ArchiveMsgFolderRoot does not work (as expected per documentation)");
        }
        Console.WriteLine(new string('=', 70));

        return successfullyAccessed;
    }
}
