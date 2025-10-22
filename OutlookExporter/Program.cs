using Azure.Identity;
using Microsoft.Extensions.Configuration;
using Microsoft.Graph;
using System.Text.Json;
using OutlookExporter;  // For EwsArchiveService

Console.WriteLine("Outlook Email Exporter");
Console.WriteLine("======================\n");

// Parse command-line arguments
string? argMailbox = null;
string? argFolder = null;
int? argCount = null;
string? argFormat = null;
bool testArchiveGuid = false;
string? testArchiveGuidValue = null;
bool useEws = false;  // Force EWS usage (for archives)

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
    else if ((args[i] == "--count" || args[i] == "-c") && i + 1 < args.Length)
    {
        if (int.TryParse(args[i + 1], out int count))
        {
            argCount = count;
        }
        else
        {
            Console.WriteLine($"Error: Invalid count value '{args[i + 1]}'. Must be a number.");
            return;
        }
        i++; // Skip next argument
    }
    else if ((args[i] == "--format" || args[i] == "-o") && i + 1 < args.Length)
    {
        string format = args[i + 1].ToLower();
        if (format == "json" || format == "html" || format == "both")
        {
            argFormat = format;
        }
        else
        {
            Console.WriteLine($"Error: Invalid format '{args[i + 1]}'. Must be 'json', 'html', or 'both'.");
            return;
        }
        i++; // Skip next argument
    }
    else if (args[i] == "--test-archive" && i + 1 < args.Length)
    {
        testArchiveGuid = true;
        testArchiveGuidValue = args[i + 1];
        i++; // Skip next argument
    }
    else if (args[i] == "--use-ews")
    {
        useEws = true;
    }
    else if (args[i] == "--help" || args[i] == "-h")
    {
        Console.WriteLine("Usage: OutlookExporter [options]");
        Console.WriteLine("\nOptions:");
        Console.WriteLine("  -m, --mailbox <email>      Specify mailbox email address");
        Console.WriteLine("  -f, --folder <name>        Specify folder name to export");
        Console.WriteLine("  -c, --count <number>       Number of emails to export (default: 5, use 0 for all)");
        Console.WriteLine("  -o, --format <format>      Output format: json, html, or both (default: json)");
        Console.WriteLine("  --use-ews                  Use EWS instead of Graph API (for archives)");
        Console.WriteLine("  --test-archive <guid>      Test if ArchiveGuid can access archive mailbox");
        Console.WriteLine("  -h, --help                 Show this help message");
        Console.WriteLine("\nExamples:");
        Console.WriteLine("  OutlookExporter --mailbox user@example.com --folder \"Sent Items\"");
        Console.WriteLine("  OutlookExporter -m user@example.com -f Inbox -c 100");
        Console.WriteLine("  OutlookExporter -m user@example.com -f Inbox -c 0  # Export all emails");
        Console.WriteLine("  OutlookExporter -m user@example.com -f Inbox -o html  # Export to HTML");
        Console.WriteLine("  OutlookExporter -m user@example.com -f Inbox -o both  # Export to JSON and HTML");
        Console.WriteLine("\nArchive mailbox access (using EWS):");
        Console.WriteLine("  OutlookExporter --use-ews -m archive@example.com -f \"Sent Items\"");
        Console.WriteLine("  OutlookExporter --use-ews -m user@example.com -f Inbox -c 100");
        Console.WriteLine("\nTesting archive access:");
        Console.WriteLine("  OutlookExporter --test-archive <archive-guid> -m <primary-email>");
        Console.WriteLine("  Example: OutlookExporter --test-archive d88b8107-7f31-4e0a-999b-723a8ba54ac0 -m Grupo.ComDev@samsys.pt");
        return;
    }
}

if (argMailbox != null)
{
    Console.WriteLine($"Command-line argument: Mailbox = {argMailbox}");
}
if (argFolder != null)
{
    Console.WriteLine($"Command-line argument: Folder = {argFolder}");
}
if (argCount != null)
{
    Console.WriteLine($"Command-line argument: Count = {(argCount == 0 ? "all" : argCount.ToString())}");
}
if (argFormat != null)
{
    Console.WriteLine($"Command-line argument: Format = {argFormat}");
}
if (useEws)
{
    Console.WriteLine("Mode: EWS (Exchange Web Services) - Archive mailbox access");
}
Console.WriteLine();

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
    // Scopes for Graph API + EWS (for archive access)
    var scopes = new[] {
        "User.Read",
        "User.Read.All",
        "Mail.Read",
        "Mail.ReadBasic",
        "Mail.Read.Shared",
        "MailboxSettings.Read",
        "https://outlook.office365.com/EWS.AccessAsUser.All"  // For EWS archive access
    };

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

    // If --test-archive flag is set, run archive GUID tests and exit
    if (testArchiveGuid && !string.IsNullOrEmpty(testArchiveGuidValue))
    {
        string primarySmtp = argMailbox ?? user?.Mail ?? user?.UserPrincipalName ?? "";

        Console.WriteLine("\n" + new string('=', 70));
        Console.WriteLine("EXPERIMENTAL: Testing ArchiveGuid Access");
        Console.WriteLine(new string('=', 70));
        Console.WriteLine("This will test if Microsoft Graph API accepts ArchiveGuid as a valid");
        Console.WriteLine("user identifier for accessing archive mailboxes.");
        Console.WriteLine("This approach is UNDOCUMENTED and may not work.");
        Console.WriteLine(new string('=', 70));

        // Run ArchiveGuid test
        bool archiveGuidWorks = await OutlookExporter.ArchiveGuidTester.TestArchiveGuidAccess(
            graphClient,
            testArchiveGuidValue,
            "Test Archive Mailbox",
            primarySmtp
        );

        // Also test ArchiveMsgFolderRoot approach
        if (!string.IsNullOrEmpty(primarySmtp))
        {
            bool archiveMsgFolderRootWorks = await OutlookExporter.ArchiveGuidTester.TestArchiveMsgFolderRoot(
                graphClient,
                primarySmtp
            );

            // Final summary
            Console.WriteLine("\n" + new string('=', 70));
            Console.WriteLine("TEST SUMMARY");
            Console.WriteLine(new string('=', 70));
            Console.WriteLine($"ArchiveGuid approach:          {(archiveGuidWorks ? "✓ WORKS!" : "✗ Does not work")}");
            Console.WriteLine($"ArchiveMsgFolderRoot approach: {(archiveMsgFolderRootWorks ? "✓ WORKS!" : "✗ Does not work")}");
            Console.WriteLine(new string('=', 70));

            if (archiveGuidWorks || archiveMsgFolderRootWorks)
            {
                Console.WriteLine("\n🎉 BREAKTHROUGH DISCOVERY! Archive access via Graph API is possible!");
                Console.WriteLine("Please report your findings and update the documentation.");
            }
            else
            {
                Console.WriteLine("\n❌ Both approaches failed. This confirms Microsoft's limitation:");
                Console.WriteLine("   Graph API does not support In-Place Archive mailbox access.");
                Console.WriteLine("   Recommendation: Use Exchange Online PowerShell or EWS instead.");
            }
        }

        Console.WriteLine("\nTest completed. Exiting.\n");
        return;
    }

    // If EWS mode is enabled, use EWS service for archive access
    if (useEws)
    {
        Console.WriteLine("\n" + new string('=', 70));
        Console.WriteLine("EWS MODE: Using Exchange Web Services for Archive Access");
        Console.WriteLine(new string('=', 70));
        Console.WriteLine("This mode accesses archive mailboxes using EWS.");
        Console.WriteLine("Graph API does not support archive mailbox access.");
        Console.WriteLine(new string('=', 70));

        if (string.IsNullOrEmpty(argMailbox))
        {
            Console.WriteLine("\nError: --use-ews requires specifying a mailbox with -m/--mailbox");
            Console.WriteLine("Example: dotnet run -- --use-ews -m archive@example.com -f \"Sent Items\"");
            return;
        }

        // Create EWS service
        var ewsService = new EwsArchiveService(clientId, tenantId);

        try
        {
            // Get archive folders
            Console.WriteLine($"\nRetrieving folders from archive mailbox: {argMailbox}");
            var archiveFolders = await ewsService.GetArchiveFoldersAsync(argMailbox);

            if (archiveFolders.Count == 0)
            {
                Console.WriteLine("\nNo folders found in archive. Make sure In-Place Archive is enabled.");
                return;
            }

            // Find or select folder
            string? selectedFolderId = null;
            string? selectedFolderName = null;

            if (!string.IsNullOrEmpty(argFolder))
            {
                // Find folder by name or path
                var matchedFolder = archiveFolders.FirstOrDefault(f =>
                    f.DisplayName.Equals(argFolder, StringComparison.OrdinalIgnoreCase) ||
                    f.Path.Equals(argFolder, StringComparison.OrdinalIgnoreCase));

                if (matchedFolder != default)
                {
                    selectedFolderId = matchedFolder.Id;
                    selectedFolderName = matchedFolder.Path;
                    Console.WriteLine($"✓ Found folder: {selectedFolderName}");
                }
                else
                {
                    Console.WriteLine($"✗ Error: Folder '{argFolder}' not found in archive.");
                    Console.WriteLine("\nAvailable folders:");
                    foreach (var folder in archiveFolders.Take(10))
                    {
                        Console.WriteLine($"  - {folder.Path}");
                    }
                    if (archiveFolders.Count > 10)
                    {
                        Console.WriteLine($"  ... and {archiveFolders.Count - 10} more folders");
                    }
                    return;
                }
            }
            else
            {
                // Interactive folder selection
                Console.WriteLine($"\nFound {archiveFolders.Count} folder(s) in archive:\n");
                for (int i = 0; i < archiveFolders.Count; i++)
                {
                    var folder = archiveFolders[i];
                    Console.WriteLine($"  [{i + 1}] {folder.Path}");
                    Console.WriteLine($"      Total Items: {folder.TotalItems}");
                    Console.WriteLine($"      Unread Items: {folder.UnreadItems}");
                    Console.WriteLine();
                }

                Console.Write("\nSelect folder to export (enter number): ");
                var folderSelection = Console.ReadLine();

                if (int.TryParse(folderSelection, out int folderIndex) &&
                    folderIndex > 0 && folderIndex <= archiveFolders.Count)
                {
                    var selectedFolder = archiveFolders[folderIndex - 1];
                    selectedFolderId = selectedFolder.Id;
                    selectedFolderName = selectedFolder.Path;
                }
                else
                {
                    Console.WriteLine("Invalid selection. Exiting.");
                    return;
                }
            }

            // Export emails using EWS
            int emailCount = argCount ?? 5;
            Console.WriteLine($"\nExporting {(emailCount == 0 ? "all" : emailCount.ToString())} email(s) from: {selectedFolderName}");

            var emails = await ewsService.ExportArchiveEmailsAsync(argMailbox, selectedFolderId!, emailCount);

            if (emails.Count > 0)
            {
                // Export to JSON (same format as Graph API export)
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
                };
                var json = JsonSerializer.Serialize(emails, jsonOptions);

                var sanitizedFolderName = string.Concat(selectedFolderName.Split(Path.GetInvalidFileNameChars()));
                var jsonOutputFile = $"exported_emails_{sanitizedFolderName}_ews.json";
                await File.WriteAllTextAsync(jsonOutputFile, json);

                Console.WriteLine($"\n✓ Exported {emails.Count} email(s) to: {jsonOutputFile}");
                Console.WriteLine($"  File size: {new FileInfo(jsonOutputFile).Length / 1024.0:F2} KB");
            }
            else
            {
                Console.WriteLine("\nNo emails found in the selected folder.");
            }

            Console.WriteLine("\nEWS export completed successfully.");
            return;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"\n✗ EWS Error: {ex.Message}");
            if (ex.InnerException != null)
            {
                Console.WriteLine($"   Inner Error: {ex.InnerException.Message}");
            }
            return;
        }
    }

    // Variables to track current export settings across multiple export cycles
    string? currentMailbox = argMailbox;
    string? currentFolder = argFolder;
    int? currentCount = argCount;
    string? currentFormat = argFormat;
    bool continueExporting = true;
    int exportCycle = 0;

    // Function to prompt for parameters with defaults
    void PromptForParameters()
    {
        Console.WriteLine("\n" + new string('=', 50));
        Console.WriteLine("Enter export parameters (press Enter to keep current value):");
        Console.WriteLine(new string('=', 50));

        // Prompt for mailbox
        Console.Write($"Mailbox [{currentMailbox ?? "none"}]: ");
        string? mailboxInput = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(mailboxInput))
        {
            currentMailbox = mailboxInput.Trim();
        }

        // Prompt for folder
        Console.Write($"Folder [{currentFolder ?? "none"}]: ");
        string? folderInput = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(folderInput))
        {
            currentFolder = folderInput.Trim();
        }

        // Prompt for count
        Console.Write($"Count [{currentCount?.ToString() ?? "5"}]: ");
        string? countInput = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(countInput))
        {
            if (int.TryParse(countInput.Trim(), out int count))
            {
                currentCount = count;
            }
            else
            {
                Console.WriteLine("Invalid count, keeping previous value.");
            }
        }

        // Prompt for format
        Console.Write($"Format [{currentFormat ?? "json"}]: ");
        string? formatInput = Console.ReadLine();
        if (!string.IsNullOrWhiteSpace(formatInput))
        {
            string format = formatInput.Trim().ToLower();
            if (format == "json" || format == "html" || format == "both")
            {
                currentFormat = format;
            }
            else
            {
                Console.WriteLine("Invalid format (must be json, html, or both), keeping previous value.");
            }
        }

        Console.WriteLine();
    }

    // Main export loop - allows multiple exports without re-authentication
    do
    {
        exportCycle++;

        // Show export cycle number if this is not the first iteration
        if (exportCycle > 1)
        {
            Console.WriteLine("\n" + new string('=', 50));
            Console.WriteLine($"Export Cycle #{exportCycle}");
            Console.WriteLine(new string('=', 50));

            // Prompt for new parameters with defaults from previous export
            PromptForParameters();
        }

        // Track previous mailbox to detect changes
        string? previousMailboxEmail = exportCycle == 1 ? null : currentMailbox;

        // Mailbox discovery is only needed if:
        // 1. First iteration AND no mailbox specified in args, OR
        // 2. Subsequent iteration AND mailbox changed from previous
        bool needMailboxDiscovery = (exportCycle == 1 && currentMailbox == null);

    var availableMailboxes = new List<(string DisplayName, string Email, string Type)>();

    if (needMailboxDiscovery)
    {
        // Discover available mailboxes
        Console.WriteLine("\n" + new string('=', 50));
        Console.WriteLine("Discovering available mailboxes...");
        Console.WriteLine(new string('=', 50));

    // Add primary mailbox
    availableMailboxes.Add((user?.DisplayName ?? "Primary Mailbox", user?.Mail ?? user?.UserPrincipalName ?? "", "Primary"));

    // Add known mailboxes from configuration
    var knownMailboxesSection = configuration.GetSection("KnownMailboxes");
    if (knownMailboxesSection.Exists())
    {
        var knownMailboxes = knownMailboxesSection.Get<List<KnownMailbox>>();
        if (knownMailboxes != null && knownMailboxes.Count > 0)
        {
            Console.WriteLine($"Adding {knownMailboxes.Count} known mailbox(es) from configuration...");
            foreach (var mailbox in knownMailboxes)
            {
                if (!string.IsNullOrEmpty(mailbox.Email))
                {
                    // Use IsArchive flag to determine mailbox type
                    string mailboxType = mailbox.IsArchive ? "Known (Archive)" : "Known";
                    string displayName = mailbox.IsArchive && !mailbox.DisplayName?.Contains("Archive") == true
                        ? $"{mailbox.DisplayName} (Archive)"
                        : mailbox.DisplayName ?? mailbox.Email;

                    availableMailboxes.Add((displayName, mailbox.Email, mailboxType));
                }
            }
        }
    }

    // Only discover mailboxes if not specified
    if (currentMailbox == null)
    {
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

        // Discover archive mailboxes for accessible mailboxes
        Console.WriteLine("\nAttempting to discover archive mailboxes...");
        Console.WriteLine("Note: Graph API does not officially support archive detection.");
        Console.WriteLine("Using heuristic patterns - results may vary by organization.\n");
        var archiveCount = 0;
        var mailboxesToCheck = new List<(string DisplayName, string Email, string Type)>(availableMailboxes);

        foreach (var mailbox in mailboxesToCheck)
        {
            try
            {
                Console.Write($"  Checking for archive: {mailbox.DisplayName}...");

                // First, try to check mailboxSettings.archiveFolder (though it's for primary mailbox archive folder, not In-Place Archive)
                try
                {
                    var settings = await graphClient.Users[mailbox.Email].MailboxSettings.GetAsync();
                    if (!string.IsNullOrEmpty(settings?.ArchiveFolder))
                    {
                        Console.Write($" [archiveFolder: {settings.ArchiveFolder.Substring(0, Math.Min(8, settings.ArchiveFolder.Length))}...]");
                    }
                }
                catch
                {
                    // MailboxSettings check failed - continue with pattern matching
                }

                // Try multiple archive naming patterns
                // Note: These are heuristic patterns and may not work for all organizations
                string localPart = mailbox.Email.Split('@')[0];
                string domain = mailbox.Email.Split('@')[1];

                string[] archivePatterns = new[]
                {
                    $"{localPart}-archive@{domain}",           // Standard pattern: user-archive@domain.com
                    $"{localPart}-Archive@{domain}",           // Capital A variant
                    $"{localPart}.archive@{domain}",           // Dot separator: user.archive@domain.com
                    $"{localPart}_archive@{domain}",           // Underscore separator: user_archive@domain.com
                    $"archive-{localPart}@{domain}",           // Prefix with dash: archive-user@domain.com
                    $"archive.{localPart}@{domain}",           // Prefix with dot: archive.user@domain.com
                    $"{localPart}@archive.{domain}",           // Subdomain: user@archive.domain.com
                    $"{localPart}-ArchiveMailbox@{domain}",    // Full word variant
                    $"{localPart}.ArchiveMailbox@{domain}",    // Full word with dot
                };

                bool found = false;
                foreach (var archiveEmail in archivePatterns)
                {
                    try
                    {
                        var archiveTest = await graphClient.Users[archiveEmail].MailFolders.GetAsync(requestConfig =>
                        {
                            requestConfig.QueryParameters.Top = 1;
                        });

                        if (archiveTest?.Value != null)
                        {
                            availableMailboxes.Add(($"{mailbox.DisplayName} (Archive)", archiveEmail, "Archive"));
                            archiveCount++;
                            Console.WriteLine($" ✓ Archive found: {archiveEmail}");
                            found = true;
                            break;
                        }
                    }
                    catch
                    {
                        // Try next pattern
                        continue;
                    }
                }

                if (!found)
                {
                    Console.WriteLine(" ✗ No archive");
                }
            }
            catch
            {
                Console.WriteLine(" ✗ Error checking");
            }
        }

        if (archiveCount > 0)
        {
            Console.WriteLine($"\nDiscovered {archiveCount} archive mailbox(es).");
        }
        else
        {
            Console.WriteLine("\nNo archive mailboxes found.");
            Console.WriteLine("Note: If you have an Online Archive, you may need to access it manually.");
            Console.WriteLine("Try using option [0] to enter the archive email address directly.");
        }
    }
    else
    {
        Console.WriteLine("\nSkipping mailbox discovery (mailbox specified).");
    }
    }
    else
    {
        Console.WriteLine("\nSkipping mailbox discovery (using previous mailbox).");
    }

    string selectedMailboxEmail = "";
    string selectedMailboxName = "";
    string? selection = null;

    // Check if mailbox was specified (from args or current setting)
    if (currentMailbox != null)
    {
        Console.WriteLine($"\nUsing mailbox: {currentMailbox}");
        selectedMailboxEmail = currentMailbox;

        // Try to find the display name from available mailboxes
        var matchedMailbox = availableMailboxes.FirstOrDefault(m =>
            m.Email.Equals(currentMailbox, StringComparison.OrdinalIgnoreCase));

        if (matchedMailbox != default)
        {
            selectedMailboxName = matchedMailbox.DisplayName;
        }
        else
        {
            selectedMailboxName = argMailbox;
        }
    }
    else
    {
        // Display available mailboxes for interactive selection
        Console.WriteLine($"\nFound {availableMailboxes.Count} mailbox(es):");
        for (int i = 0; i < availableMailboxes.Count; i++)
        {
            Console.WriteLine($"  [{i + 1}] {availableMailboxes[i].DisplayName} ({availableMailboxes[i].Email}) - {availableMailboxes[i].Type}");
        }
        Console.WriteLine($"  [0] Enter custom mailbox email address");

        Console.Write("\nSelect mailbox (enter number): ");
        selection = Console.ReadLine();
    }

    if (selection != null && int.TryParse(selection, out int selectedIndex))
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
    else if (selection != null)
    {
        Console.WriteLine("Invalid input, using primary mailbox.");
        selectedMailboxEmail = user?.Mail ?? user?.UserPrincipalName ?? "";
        selectedMailboxName = user?.DisplayName ?? "Primary";
    }

    Console.WriteLine($"\nSelected mailbox: {selectedMailboxName} ({selectedMailboxEmail})");

    // List all mail folders
    Console.WriteLine("\n" + new string('=', 50));
    Console.WriteLine("Retrieving mail folders (including subfolders)...");
    Console.WriteLine(new string('=', 50));

    // Get all folders including nested subfolders
    var allFolders = new List<(string Id, string DisplayName, string Path, int TotalItems, int UnreadItems)>();
    bool folderFound = false; // Flag to stop searching when target folder is found

    async Task GetFoldersRecursive(string parentFolderId, string parentPath)
    {
        var childFolders = await graphClient.Users[selectedMailboxEmail]
            .MailFolders[parentFolderId]
            .ChildFolders
            .GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Top = 999; // Maximum per page
            });

        if (childFolders?.Value != null)
        {
            // Use PageIterator to handle all pages of child folders
            var pageIterator = Microsoft.Graph.PageIterator<Microsoft.Graph.Models.MailFolder, Microsoft.Graph.Models.MailFolderCollectionResponse>
                .CreatePageIterator(graphClient, childFolders, (folder) =>
                {
                    if (folderFound) return false; // Stop if we already found the target folder

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

                    // Check if this is the target folder (if specified via CLI)
                    if (currentFolder != null &&
                        (folderPath.Equals(currentFolder, StringComparison.OrdinalIgnoreCase) ||
                         folder.DisplayName?.Equals(currentFolder, StringComparison.OrdinalIgnoreCase) == true))
                    {
                        folderFound = true;
                        return false; // Stop searching
                    }

                    // Recursively get child folders
                    if (folder.ChildFolderCount > 0 && !folderFound)
                    {
                        GetFoldersRecursive(folder.Id ?? "", folderPath).Wait();
                    }

                    return !folderFound; // Continue iterating only if not found
                });

            await pageIterator.IterateAsync();
        }
    }

    // Get root level folders first with pagination support
    var rootFolders = await graphClient.Users[selectedMailboxEmail].MailFolders.GetAsync(requestConfig =>
    {
        requestConfig.QueryParameters.Top = 999; // Maximum per page
    });

    if (rootFolders?.Value != null)
    {
        // Use PageIterator to handle all pages of root folders
        var rootPageIterator = Microsoft.Graph.PageIterator<Microsoft.Graph.Models.MailFolder, Microsoft.Graph.Models.MailFolderCollectionResponse>
            .CreatePageIterator(graphClient, rootFolders, (folder) =>
            {
                if (folderFound) return false; // Stop if we already found the target folder

                allFolders.Add((
                    folder.Id ?? "",
                    folder.DisplayName ?? "",
                    folder.DisplayName ?? "",
                    folder.TotalItemCount ?? 0,
                    folder.UnreadItemCount ?? 0
                ));

                // Check if this is the target folder (if specified via CLI)
                if (currentFolder != null &&
                    folder.DisplayName?.Equals(currentFolder, StringComparison.OrdinalIgnoreCase) == true)
                {
                    folderFound = true;
                    return false; // Stop searching
                }

                // Get subfolders if any
                if (folder.ChildFolderCount > 0 && !folderFound)
                {
                    GetFoldersRecursive(folder.Id ?? "", folder.DisplayName ?? "").Wait();
                }

                return !folderFound; // Continue iterating only if not found
            });

        await rootPageIterator.IterateAsync();
    }

    string? selectedFolderId = null;
    string selectedFolderName = "Inbox";

    if (allFolders.Count > 0)
    {
        // Only print all folders if no specific folder was requested or if we're in interactive mode
        if (currentFolder == null)
        {
            Console.WriteLine($"\nFound {allFolders.Count} mail folder(s) (including subfolders):\n");

            for (int i = 0; i < allFolders.Count; i++)
            {
                var folder = allFolders[i];
                Console.WriteLine($"  [{i + 1}] {folder.Path}");
                Console.WriteLine($"      Total Items: {folder.TotalItems}");
                Console.WriteLine($"      Unread Items: {folder.UnreadItems}");
                Console.WriteLine();
            }
        }
        else
        {
            Console.WriteLine($"\nFound {allFolders.Count} mail folder(s) (stopped early - target folder found).");
        }

        Console.WriteLine("\nFolder listing completed successfully.");

        // Check if folder was provided via command-line argument
        if (currentFolder != null)
        {
            Console.WriteLine($"\nUsing folder from command-line argument: {currentFolder}");

            // Try to find folder by name or path (case-insensitive)
            var matchedFolder = allFolders.FirstOrDefault(f =>
                f.DisplayName.Equals(currentFolder, StringComparison.OrdinalIgnoreCase) ||
                f.Path.Equals(currentFolder, StringComparison.OrdinalIgnoreCase));

            if (matchedFolder != default)
            {
                selectedFolderId = matchedFolder.Id;
                selectedFolderName = matchedFolder.Path;
                Console.WriteLine($"✓ Found folder: {selectedFolderName}");
            }
            else
            {
                Console.WriteLine($"✗ Error: Folder '{currentFolder}' not found.");
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
                return;
            }
        }
        else
        {
            // Ask user which folder to export
            Console.Write("\nSelect folder to export (enter number, or press Enter for Inbox): ");
            var folderSelection = Console.ReadLine();

            if (!string.IsNullOrWhiteSpace(folderSelection) && int.TryParse(folderSelection, out int folderIndex))
            {
                if (folderIndex > 0 && folderIndex <= allFolders.Count)
                {
                    var selectedFolder = allFolders[folderIndex - 1];
                    selectedFolderId = selectedFolder.Id;
                    selectedFolderName = selectedFolder.Path;
                }
                else
                {
                    Console.WriteLine("Invalid selection, using Inbox.");
                    var inboxFolder = allFolders.FirstOrDefault(f => f.DisplayName.Equals("Inbox", StringComparison.OrdinalIgnoreCase));
                    selectedFolderId = inboxFolder.Id;
                    selectedFolderName = "Inbox";
                }
            }
            else
            {
                // Default to Inbox
                var inboxFolder = allFolders.FirstOrDefault(f => f.DisplayName.Equals("Inbox", StringComparison.OrdinalIgnoreCase));
                selectedFolderId = inboxFolder.Id;
                selectedFolderName = "Inbox";
            }
        }
    }
    else
    {
        Console.WriteLine("\nNo folders found.");
    }

    // Prompt for export count and format on first cycle if not provided via CLI
    if (exportCycle == 1)
    {
        if (currentCount == null)
        {
            Console.WriteLine("\n" + new string('=', 50));
            Console.Write("How many emails to export? (press Enter for 5, or 0 for all): ");
            string? countInput = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(countInput))
            {
                if (int.TryParse(countInput.Trim(), out int count))
                {
                    currentCount = count;
                }
                else
                {
                    Console.WriteLine("Invalid count, using default (5).");
                    currentCount = 5;
                }
            }
            else
            {
                currentCount = 5; // Default
            }
        }

        if (currentFormat == null)
        {
            Console.Write("Export format? (json/html/both, press Enter for json): ");
            string? formatInput = Console.ReadLine();
            if (!string.IsNullOrWhiteSpace(formatInput))
            {
                string format = formatInput.Trim().ToLower();
                if (format == "json" || format == "html" || format == "both")
                {
                    currentFormat = format;
                }
                else
                {
                    Console.WriteLine("Invalid format, using json.");
                    currentFormat = "json";
                }
            }
            else
            {
                currentFormat = "json"; // Default
            }
            Console.WriteLine(new string('=', 50));
        }
    }

    // HTML Export Function
    string GenerateHtmlExport(List<Microsoft.Graph.Models.Message> messages, string folderName, string mailboxEmail)
    {
        var html = new System.Text.StringBuilder();

        html.AppendLine("<!DOCTYPE html>");
        html.AppendLine("<html lang=\"en\">");
        html.AppendLine("<head>");
        html.AppendLine("    <meta charset=\"UTF-8\">");
        html.AppendLine("    <meta name=\"viewport\" content=\"width=device-width, initial-scale=1.0\">");
        html.AppendLine($"    <title>Email Export - {folderName}</title>");
        html.AppendLine("    <style>");
        html.AppendLine("        * { margin: 0; padding: 0; box-sizing: border-box; }");
        html.AppendLine("        body { font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, 'Helvetica Neue', Arial, sans-serif; background-color: #f5f5f5; color: #333; line-height: 1.6; padding: 20px; }");
        html.AppendLine("        .container { max-width: 1200px; margin: 0 auto; background-color: white; box-shadow: 0 2px 8px rgba(0,0,0,0.1); }");
        html.AppendLine("        .header { background: linear-gradient(135deg, #667eea 0%, #764ba2 100%); color: white; padding: 30px; text-align: center; }");
        html.AppendLine("        .header h1 { font-size: 28px; margin-bottom: 10px; }");
        html.AppendLine("        .header .subtitle { font-size: 16px; opacity: 0.9; }");
        html.AppendLine("        .header .export-info { margin-top: 15px; font-size: 14px; opacity: 0.8; }");
        html.AppendLine("        .email-card { border-bottom: 3px solid #f0f0f0; padding: 30px; background-color: white; }");
        html.AppendLine("        .email-card:nth-child(even) { background-color: #fafafa; }");
        html.AppendLine("        .email-number { display: inline-block; background-color: #667eea; color: white; padding: 4px 12px; border-radius: 12px; font-size: 12px; font-weight: bold; margin-bottom: 15px; }");
        html.AppendLine("        .metadata-table { width: 100%; border-collapse: collapse; margin-bottom: 20px; background-color: white; border: 1px solid #e0e0e0; }");
        html.AppendLine("        .metadata-table th { background-color: #f8f9fa; text-align: left; padding: 12px; font-weight: 600; color: #495057; border-bottom: 2px solid #dee2e6; width: 180px; }");
        html.AppendLine("        .metadata-table td { padding: 12px; border-bottom: 1px solid #e9ecef; }");
        html.AppendLine("        .metadata-table tr:last-child td { border-bottom: none; }");
        html.AppendLine("        .email-address { font-family: 'Courier New', monospace; background-color: #e7f3ff; padding: 2px 6px; border-radius: 3px; font-size: 13px; }");
        html.AppendLine("        .subject { font-size: 20px; font-weight: 600; color: #2c3e50; margin-bottom: 15px; }");
        html.AppendLine("        .email-body { background-color: #f9f9f9; border-left: 4px solid #667eea; padding: 20px; margin-top: 15px; border-radius: 4px; max-height: 500px; overflow-y: auto; }");
        html.AppendLine("        .email-body-content { line-height: 1.8; color: #555; }");
        html.AppendLine("        .badge { display: inline-block; padding: 3px 8px; border-radius: 3px; font-size: 11px; font-weight: 600; text-transform: uppercase; margin-right: 5px; }");
        html.AppendLine("        .badge-read { background-color: #d4edda; color: #155724; }");
        html.AppendLine("        .badge-unread { background-color: #fff3cd; color: #856404; }");
        html.AppendLine("        .badge-important { background-color: #f8d7da; color: #721c24; }");
        html.AppendLine("        .badge-draft { background-color: #d1ecf1; color: #0c5460; }");
        html.AppendLine("        .recipients { display: block; margin-top: 5px; }");
        html.AppendLine("        .recipient-item { display: inline-block; margin-right: 10px; margin-bottom: 5px; }");
        html.AppendLine("        @media print { body { background-color: white; } .container { box-shadow: none; } .email-card { page-break-inside: avoid; } }");
        html.AppendLine("        @media (max-width: 768px) { .container { margin: 0; } .header { padding: 20px; } .email-card { padding: 20px; } .metadata-table th { width: 120px; font-size: 13px; } }");
        html.AppendLine("    </style>");
        html.AppendLine("</head>");
        html.AppendLine("<body>");
        html.AppendLine("    <div class=\"container\">");
        html.AppendLine("        <div class=\"header\">");
        html.AppendLine($"            <h1>Email Export</h1>");
        html.AppendLine($"            <div class=\"subtitle\">Folder: {System.Web.HttpUtility.HtmlEncode(folderName)} | Mailbox: {System.Web.HttpUtility.HtmlEncode(mailboxEmail)}</div>");
        html.AppendLine($"            <div class=\"export-info\">Total Emails: {messages.Count} | Exported: {DateTime.Now:yyyy-MM-dd HH:mm:ss}</div>");
        html.AppendLine("        </div>");

        for (int i = 0; i < messages.Count; i++)
        {
            var msg = messages[i];
            html.AppendLine($"        <div class=\"email-card\">");
            html.AppendLine($"            <span class=\"email-number\">Email #{i + 1}</span>");

            // Subject
            html.AppendLine($"            <div class=\"subject\">{System.Web.HttpUtility.HtmlEncode(msg.Subject ?? "(No Subject)")}</div>");

            // Badges
            if (msg.IsRead == true)
                html.AppendLine("            <span class=\"badge badge-read\">Read</span>");
            else
                html.AppendLine("            <span class=\"badge badge-unread\">Unread</span>");

            if (msg.Importance?.ToString() == "High")
                html.AppendLine("            <span class=\"badge badge-important\">Important</span>");

            if (msg.IsDraft == true)
                html.AppendLine("            <span class=\"badge badge-draft\">Draft</span>");

            // Metadata Table
            html.AppendLine("            <table class=\"metadata-table\">");

            // From
            html.AppendLine("                <tr>");
            html.AppendLine("                    <th>From</th>");
            html.AppendLine($"                    <td><span class=\"email-address\">{System.Web.HttpUtility.HtmlEncode(msg.From?.EmailAddress?.Address ?? "Unknown")}</span> ({System.Web.HttpUtility.HtmlEncode(msg.From?.EmailAddress?.Name ?? "Unknown")})</td>");
            html.AppendLine("                </tr>");

            // To
            if (msg.ToRecipients?.Any() == true)
            {
                html.AppendLine("                <tr>");
                html.AppendLine("                    <th>To</th>");
                html.AppendLine("                    <td><div class=\"recipients\">");
                foreach (var recipient in msg.ToRecipients)
                {
                    html.AppendLine($"                        <span class=\"recipient-item\"><span class=\"email-address\">{System.Web.HttpUtility.HtmlEncode(recipient.EmailAddress?.Address ?? "")}</span> ({System.Web.HttpUtility.HtmlEncode(recipient.EmailAddress?.Name ?? "")})</span>");
                }
                html.AppendLine("                    </div></td>");
                html.AppendLine("                </tr>");
            }

            // Cc
            if (msg.CcRecipients?.Any() == true)
            {
                html.AppendLine("                <tr>");
                html.AppendLine("                    <th>Cc</th>");
                html.AppendLine("                    <td><div class=\"recipients\">");
                foreach (var recipient in msg.CcRecipients)
                {
                    html.AppendLine($"                        <span class=\"recipient-item\"><span class=\"email-address\">{System.Web.HttpUtility.HtmlEncode(recipient.EmailAddress?.Address ?? "")}</span> ({System.Web.HttpUtility.HtmlEncode(recipient.EmailAddress?.Name ?? "")})</span>");
                }
                html.AppendLine("                    </div></td>");
                html.AppendLine("                </tr>");
            }

            // Bcc
            if (msg.BccRecipients?.Any() == true)
            {
                html.AppendLine("                <tr>");
                html.AppendLine("                    <th>Bcc</th>");
                html.AppendLine("                    <td><div class=\"recipients\">");
                foreach (var recipient in msg.BccRecipients)
                {
                    html.AppendLine($"                        <span class=\"recipient-item\"><span class=\"email-address\">{System.Web.HttpUtility.HtmlEncode(recipient.EmailAddress?.Address ?? "")}</span> ({System.Web.HttpUtility.HtmlEncode(recipient.EmailAddress?.Name ?? "")})</span>");
                }
                html.AppendLine("                    </div></td>");
                html.AppendLine("                </tr>");
            }

            // Received Date
            html.AppendLine("                <tr>");
            html.AppendLine("                    <th>Received</th>");
            html.AppendLine($"                    <td>{msg.ReceivedDateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Unknown"}</td>");
            html.AppendLine("                </tr>");

            // Sent Date
            html.AppendLine("                <tr>");
            html.AppendLine("                    <th>Sent</th>");
            html.AppendLine($"                    <td>{msg.SentDateTime?.ToString("yyyy-MM-dd HH:mm:ss") ?? "Unknown"}</td>");
            html.AppendLine("                </tr>");

            // Importance
            html.AppendLine("                <tr>");
            html.AppendLine("                    <th>Importance</th>");
            html.AppendLine($"                    <td>{msg.Importance?.ToString() ?? "Normal"}</td>");
            html.AppendLine("                </tr>");

            // Has Attachments
            html.AppendLine("                <tr>");
            html.AppendLine("                    <th>Has Attachments</th>");
            html.AppendLine($"                    <td>{(msg.HasAttachments == true ? "Yes" : "No")}</td>");
            html.AppendLine("                </tr>");

            // Categories
            if (msg.Categories?.Any() == true)
            {
                html.AppendLine("                <tr>");
                html.AppendLine("                    <th>Categories</th>");
                html.AppendLine($"                    <td>{System.Web.HttpUtility.HtmlEncode(string.Join(", ", msg.Categories))}</td>");
                html.AppendLine("                </tr>");
            }

            // Conversation ID
            if (!string.IsNullOrEmpty(msg.ConversationId))
            {
                html.AppendLine("                <tr>");
                html.AppendLine("                    <th>Conversation ID</th>");
                html.AppendLine($"                    <td><span class=\"email-address\">{System.Web.HttpUtility.HtmlEncode(msg.ConversationId)}</span></td>");
                html.AppendLine("                </tr>");
            }

            html.AppendLine("            </table>");

            // Email Body
            html.AppendLine("            <div class=\"email-body\">");
            html.AppendLine("                <div class=\"email-body-content\">");

            if (msg.Body?.ContentType?.ToString() == "Html" && !string.IsNullOrEmpty(msg.Body.Content))
            {
                // Render HTML body (sanitized)
                html.AppendLine($"                    {msg.Body.Content}");
            }
            else if (!string.IsNullOrEmpty(msg.Body?.Content))
            {
                // Render plain text body
                html.AppendLine($"                    <pre style=\"white-space: pre-wrap; word-wrap: break-word; font-family: inherit;\">{System.Web.HttpUtility.HtmlEncode(msg.Body.Content)}</pre>");
            }
            else if (!string.IsNullOrEmpty(msg.BodyPreview))
            {
                // Fallback to body preview
                html.AppendLine($"                    <p>{System.Web.HttpUtility.HtmlEncode(msg.BodyPreview)}</p>");
            }
            else
            {
                html.AppendLine("                    <p><em>(No content available)</em></p>");
            }

            html.AppendLine("                </div>");
            html.AppendLine("            </div>");
            html.AppendLine("        </div>");
        }

        html.AppendLine("    </div>");
        html.AppendLine("</body>");
        html.AppendLine("</html>");

        return html.ToString();
    }

    // Determine export format (default to JSON for backward compatibility)
    string exportFormat = currentFormat ?? "json";

    // Export emails
    Console.WriteLine("\n" + new string('=', 50));
    Console.WriteLine($"Exporting emails from {selectedFolderName}...");
    Console.WriteLine(new string('=', 50));

    // Determine how many emails to export
    int emailCount = currentCount ?? 5; // Default to 5 if not specified
    bool exportAll = currentCount == 0;

    var allMessages = new List<Microsoft.Graph.Models.Message>();

    if (exportAll)
    {
        Console.WriteLine("\nExporting all emails (this may take a while)...");

        // Get all emails with pagination
        var messages = await graphClient.Users[selectedMailboxEmail].MailFolders[selectedFolderId ?? "Inbox"].Messages
            .GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Top = 1000; // Maximum per page
            });

        if (messages?.Value != null)
        {
            allMessages.AddRange(messages.Value);
            Console.WriteLine($"Retrieved {allMessages.Count} emails...");

            // Handle pagination if there are more results
            var pageIterator = Microsoft.Graph.PageIterator<Microsoft.Graph.Models.Message, Microsoft.Graph.Models.MessageCollectionResponse>
                .CreatePageIterator(graphClient, messages, (msg) =>
                {
                    allMessages.Add(msg);
                    if (allMessages.Count % 1000 == 0)
                    {
                        Console.WriteLine($"Retrieved {allMessages.Count} emails...");
                    }
                    return true; // Continue iterating
                });

            await pageIterator.IterateAsync();
            Console.WriteLine($"\nTotal retrieved: {allMessages.Count} emails");
        }
    }
    else
    {
        Console.WriteLine($"\nExporting up to {emailCount} emails...");

        // Get specified number of emails
        var messages = await graphClient.Users[selectedMailboxEmail].MailFolders[selectedFolderId ?? "Inbox"].Messages
            .GetAsync(requestConfig =>
            {
                requestConfig.QueryParameters.Top = emailCount;
            });

        if (messages?.Value != null)
        {
            allMessages.AddRange(messages.Value);
        }
    }

    if (allMessages.Count > 0)
    {
        Console.WriteLine($"\nRetrieved {allMessages.Count} emails");

        // Convert to anonymous objects for JSON export (all properties except attachments)
        var emailData = allMessages.Select(msg => new
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

        // Sanitize folder name for file system
        var sanitizedFolderName = string.Concat(selectedFolderName.Split(Path.GetInvalidFileNameChars()));

        // Export to JSON format
        if (exportFormat == "json" || exportFormat == "both")
        {
            var jsonOptions = new JsonSerializerOptions
            {
                WriteIndented = true,
                Encoder = System.Text.Encodings.Web.JavaScriptEncoder.UnsafeRelaxedJsonEscaping
            };
            var json = JsonSerializer.Serialize(emailData, jsonOptions);

            var jsonOutputFile = $"exported_emails_{sanitizedFolderName}.json";
            await File.WriteAllTextAsync(jsonOutputFile, json);

            Console.WriteLine($"✓ Exported {emailData.Count} emails to JSON: {jsonOutputFile}");
            Console.WriteLine($"  File size: {new FileInfo(jsonOutputFile).Length / 1024.0:F2} KB");
        }

        // Export to HTML format
        if (exportFormat == "html" || exportFormat == "both")
        {
            var html = GenerateHtmlExport(allMessages, selectedFolderName, selectedMailboxEmail);

            var htmlOutputFile = $"exported_emails_{sanitizedFolderName}.html";
            await File.WriteAllTextAsync(htmlOutputFile, html);

            Console.WriteLine($"✓ Exported {allMessages.Count} emails to HTML: {htmlOutputFile}");
            Console.WriteLine($"  File size: {new FileInfo(htmlOutputFile).Length / 1024.0:F2} KB");
        }
    }
    else
    {
        Console.WriteLine($"\nNo emails found in {selectedFolderName}.");
    }

    Console.WriteLine("\nExport completed successfully.");

    // Ask if user wants to continue with another export
    Console.Write("\nDo you want to export another folder? (y/n): ");
    string? continueResponse = Console.ReadLine();
    continueExporting = (continueResponse?.Trim().ToLower() == "y");

    } while (continueExporting);

    Console.WriteLine("\nDone.");
}
catch (Exception ex)
{
    Console.WriteLine($"\nError: {ex.Message}");
    if (ex.InnerException != null)
    {
        Console.WriteLine($"Inner Error: {ex.InnerException.Message}");
    }
    Console.WriteLine("\nDone.");
}

// Configuration model for known mailboxes
public class KnownMailbox
{
    public string? DisplayName { get; set; }
    public string? Email { get; set; }
    public bool IsArchive { get; set; } = false;  // Flag to explicitly mark known archives
}
