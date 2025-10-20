using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System.Text.Json;

Console.WriteLine("Outlook Email Exporter");
Console.WriteLine("======================\n");

// Load configuration from appsettings.json
var configuration = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .AddJsonFile("appsettings.Development.json", optional: true, reloadOnChange: true)
    .Build();

// Read Azure AD configuration
var clientId = configuration["AzureAd:ClientId"]
    ?? throw new InvalidOperationException("ClientId not found in configuration");
var tenantId = configuration["AzureAd:TenantId"] ?? "common";

try
{
    Console.WriteLine("Initializing authentication...");

    // Create DeviceCodeCredential for authentication
    var scopes = new[] { "User.Read", "User.Read.All", "Mail.Read", "Mail.ReadBasic", "Mail.Read.Shared", "MailboxSettings.Read" };

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

    // Create Graph client with explicit scopes
    var graphClient = new GraphServiceClient(credential, scopes);

    Console.WriteLine("\nAttempting to authenticate...");

    // Test the connection by getting user profile
    var user = await graphClient.Me.GetAsync();

    Console.WriteLine($"\nAuthentication successful!");
    Console.WriteLine($"Logged in as: {user?.DisplayName}");
    Console.WriteLine($"Email: {user?.Mail ?? user?.UserPrincipalName}");

    // Discover available mailboxes
    Console.WriteLine("\n" + new string('=', 50));
    Console.WriteLine("Discovering available mailboxes...");
    Console.WriteLine(new string('=', 50));

    var availableMailboxes = new List<(string DisplayName, string Email, string Type)>();

    // Add primary mailbox
    availableMailboxes.Add((user?.DisplayName ?? "Primary Mailbox", user?.Mail ?? user?.UserPrincipalName ?? "", "Primary"));

    // Add known mailbox
    availableMailboxes.Add(("Arquivo ComDev - SAMSYS", "arquivo.comdev@samsys.pt", "Delegated"));

    // Try to discover shared mailboxes
    Console.WriteLine("\nAttempting to discover shared/delegated mailboxes...");
    try
    {
        // Query Azure AD for shared mailboxes
        // Shared mailboxes typically have accountEnabled = false and a mailbox
        Console.WriteLine("Querying Azure AD for shared mailboxes...");

        // Get all users with mailboxes that have accountEnabled = false (typical for shared mailboxes)
        var usersResponse = await graphClient.Users.GetAsync(requestConfig =>
        {
            requestConfig.QueryParameters.Filter = "accountEnabled eq false";
            requestConfig.QueryParameters.Select = new[] { "displayName", "mail", "userPrincipalName", "id" };
            requestConfig.QueryParameters.Top = 100; // Limit to avoid large queries
        });

        var sharedMailboxCount = 0;

        if (usersResponse?.Value != null && usersResponse.Value.Count > 0)
        {
            Console.WriteLine($"Found {usersResponse.Value.Count} potential shared mailbox(es). Testing access...\n");

            foreach (var potentialSharedMailbox in usersResponse.Value)
            {
                var mailboxEmail = potentialSharedMailbox.Mail ?? potentialSharedMailbox.UserPrincipalName;

                if (string.IsNullOrEmpty(mailboxEmail))
                {
                    continue;
                }

                // Test if current user has access to this shared mailbox
                try
                {
                    Console.Write($"  Testing access to: {potentialSharedMailbox.DisplayName} ({mailboxEmail})... ");

                    // Try to get the inbox to verify access
                    var testAccess = await graphClient.Users[mailboxEmail].MailFolders.GetAsync(requestConfig =>
                    {
                        requestConfig.QueryParameters.Top = 1;
                    });

                    // If we get here, we have access
                    Console.WriteLine("✓ Access granted");
                    availableMailboxes.Add((potentialSharedMailbox.DisplayName ?? mailboxEmail, mailboxEmail, "Shared"));
                    sharedMailboxCount++;
                }
                catch (Exception)
                {
                    // No access to this mailbox - silently skip
                    Console.WriteLine("✗ No access");
                }
            }

            Console.WriteLine($"\nDiscovered {sharedMailboxCount} accessible shared mailbox(es).");
        }
        else
        {
            Console.WriteLine("No shared mailboxes found in Azure AD query.");
        }
    }
    catch (Exception ex)
    {
        Console.WriteLine($"Error during shared mailbox discovery: {ex.Message}");
        Console.WriteLine("You can still manually enter shared mailbox addresses below.");
    }

    Console.WriteLine($"\nFound {availableMailboxes.Count} mailbox(es):");
    for (int i = 0; i < availableMailboxes.Count; i++)
    {
        Console.WriteLine($"  [{i + 1}] {availableMailboxes[i].DisplayName} ({availableMailboxes[i].Email}) - {availableMailboxes[i].Type}");
    }
    Console.WriteLine($"  [0] Enter custom mailbox email address");

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

            // Validate access to the mailbox
            Console.WriteLine($"\nValidating access to {selectedMailboxEmail}...");
            try
            {
                var testAccess = await graphClient.Users[selectedMailboxEmail].MailFolders.GetAsync(requestConfig =>
                {
                    requestConfig.QueryParameters.Top = 1;
                });

                Console.WriteLine("✓ Access confirmed! You have permission to access this mailbox.");
            }
            catch (Exception validateEx)
            {
                Console.WriteLine($"✗ Access validation failed: {validateEx.Message}");
                Console.WriteLine("\nPossible reasons:");
                Console.WriteLine("  1. Missing 'Mail.Read.Shared' permission in Azure Portal");
                Console.WriteLine("  2. No 'Full Access' permission on this mailbox in Exchange");
                Console.WriteLine("  3. Incorrect mailbox email address");
                Console.WriteLine("  4. Admin consent not granted (organizational accounts)");
                Console.WriteLine("\nWould you like to continue anyway? (y/n): ");
                var cont = Console.ReadLine();
                if (cont?.ToLower() != "y")
                {
                    Console.WriteLine("Returning to primary mailbox.");
                    selectedMailboxEmail = user?.Mail ?? user?.UserPrincipalName ?? "";
                    selectedMailboxName = user?.DisplayName ?? "Primary";
                }
            }
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
    else
    {
        Console.WriteLine("Invalid input, using primary mailbox.");
        selectedMailboxEmail = user?.Mail ?? user?.UserPrincipalName ?? "";
        selectedMailboxName = user?.DisplayName ?? "Primary";
    }

    Console.WriteLine($"\nSelected mailbox: {selectedMailboxName} ({selectedMailboxEmail})");

    // List all mail folders
    Console.WriteLine("\n" + new string('=', 50));
    Console.WriteLine("Retrieving mail folders...");
    Console.WriteLine(new string('=', 50));

    var folders = await graphClient.Users[selectedMailboxEmail].MailFolders.GetAsync();

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

    Console.WriteLine("\nFolder listing completed successfully.");

    // Export emails to JSON
    Console.WriteLine("\n" + new string('=', 50));
    Console.WriteLine("Exporting emails from Inbox to JSON...");
    Console.WriteLine(new string('=', 50));

    // Get emails with all properties except attachments
    var messages = await graphClient.Users[selectedMailboxEmail].MailFolders["Inbox"].Messages
        .GetAsync(requestConfig =>
        {
            requestConfig.QueryParameters.Top = 5; // Get only 5 emails for testing
        });

    if (messages?.Value != null && messages.Value.Count > 0)
    {
        Console.WriteLine($"\nRetrieved {messages.Value.Count} emails");

        // Convert to anonymous objects for JSON export (all properties except attachments)
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
            ReplyTo = msg.ReplyTo?.Select(r => new
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
            BodyPreview = msg.BodyPreview,
            Flag = new
            {
                FlagStatus = msg.Flag?.FlagStatus?.ToString()
            }
        }).ToList();

        // Serialize to JSON with nice formatting
        var jsonOptions = new JsonSerializerOptions
        {
            WriteIndented = true,
            Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
        };
        var json = JsonSerializer.Serialize(emailData, jsonOptions);

        // Save to file
        var outputFile = "exported_emails.json";
        await File.WriteAllTextAsync(outputFile, json);

        Console.WriteLine($"✓ Exported {emailData.Count} emails to: {outputFile}");
        Console.WriteLine($"  File size: {new FileInfo(outputFile).Length / 1024.0:F2} KB");
    }
    else
    {
        Console.WriteLine("\nNo emails found in Inbox.");
    }

    Console.WriteLine("\nExport completed successfully.");
}
catch (Exception ex)
{
    Console.WriteLine($"\nError: {ex.Message}");
    if (ex.InnerException != null)
    {
        Console.WriteLine($"Inner Error: {ex.InnerException.Message}");
    }
}

Console.WriteLine("\nDone.");
