using Microsoft.Exchange.WebServices.Data;
using Azure.Identity;
using SystemTask = System.Threading.Tasks.Task;

namespace OutlookExporter;

/// <summary>
/// Service for accessing Exchange Online archive mailboxes using EWS (Exchange Web Services).
/// This is necessary because Microsoft Graph API does not support archive mailbox access.
/// </summary>
public class EwsArchiveService
{
    private readonly string _clientId;
    private readonly string _tenantId;

    public EwsArchiveService(string clientId, string tenantId)
    {
        _clientId = clientId;
        _tenantId = tenantId;
    }

    /// <summary>
    /// Creates an authenticated EWS ExchangeService using Device Code Flow
    /// </summary>
    private async System.Threading.Tasks.Task<(ExchangeService Service, string AccessToken)> CreateEwsServiceAsync()
    {
        Console.WriteLine("Acquiring EWS access token...");

        var options = new DeviceCodeCredentialOptions
        {
            ClientId = _clientId,
            TenantId = _tenantId,
            DeviceCodeCallback = (code, cancellation) =>
            {
                // Reuse the same device code from initial authentication
                // The token cache should handle this
                return SystemTask.CompletedTask;
            }
        };

        var credential = new DeviceCodeCredential(options);

        // Request token with EWS scope
        var tokenRequestContext = new Azure.Core.TokenRequestContext(
            new[] { "https://outlook.office365.com/EWS.AccessAsUser.All" }
        );

        var tokenResult = await credential.GetTokenAsync(tokenRequestContext, default);
        var accessToken = tokenResult.Token;

        Console.WriteLine("✓ EWS access token acquired");

        // Create ExchangeService
        var service = new ExchangeService(ExchangeVersion.Exchange2016)
        {
            Url = new Uri("https://outlook.office365.com/EWS/Exchange.asmx"),
            Credentials = new OAuthCredentials(accessToken),
            TraceEnabled = false, // Set to true for debugging
            TraceFlags = TraceFlags.None
        };

        return (service, accessToken);
    }

    /// <summary>
    /// Gets all folders from an archive mailbox
    /// </summary>
    public async System.Threading.Tasks.Task<List<(string Id, string DisplayName, string Path, int TotalItems, int UnreadItems)>> GetArchiveFoldersAsync(
        string mailboxEmail)
    {
        Console.WriteLine($"\n[EWS] Accessing archive mailbox for: {mailboxEmail}");

        var (service, _) = await CreateEwsServiceAsync();
        var folders = new List<(string Id, string DisplayName, string Path, int TotalItems, int UnreadItems)>();

        try
        {
            // Access the archive root folder
            var archiveRoot = await SystemTask.Run(() => Folder.Bind(service, WellKnownFolderName.ArchiveMsgFolderRoot));

            Console.WriteLine($"✓ Archive root accessed: {archiveRoot.DisplayName}");
            Console.WriteLine($"  Total items in root: {archiveRoot.TotalCount}");
            Console.WriteLine($"  Child folder count: {archiveRoot.ChildFolderCount}");

            // Recursively get all folders
            await GetFoldersRecursiveAsync(service, archiveRoot.Id, "", folders);

            Console.WriteLine($"✓ Retrieved {folders.Count} folder(s) from archive");

            return folders;
        }
        catch (ServiceResponseException ex) when (ex.Message.Contains("ErrorItemNotFound") || ex.Message.Contains("not be found"))
        {
            Console.WriteLine("✗ Archive mailbox not found or In-Place Archive is not enabled");
            Console.WriteLine("  Make sure the mailbox has an active archive in Exchange Admin Center");
            throw new InvalidOperationException("Archive mailbox not accessible. Ensure In-Place Archive is enabled.", ex);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error accessing archive: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Recursively retrieves all folders from the archive
    /// </summary>
    private async SystemTask GetFoldersRecursiveAsync(
        ExchangeService service,
        FolderId parentFolderId,
        string parentPath,
        List<(string Id, string DisplayName, string Path, int TotalItems, int UnreadItems)> folders)
    {
        try
        {
            // Define folder view with pagination
            var folderView = new FolderView(1000)
            {
                Traversal = FolderTraversal.Shallow,
                PropertySet = new PropertySet(BasePropertySet.FirstClassProperties)
            };

            FindFoldersResults findResults;

            do
            {
                findResults = await SystemTask.Run(() => service.FindFolders(parentFolderId, folderView));

                foreach (var folder in findResults.Folders)
                {
                    var folderPath = string.IsNullOrEmpty(parentPath)
                        ? folder.DisplayName
                        : $"{parentPath}/{folder.DisplayName}";

                    folders.Add((
                        folder.Id.UniqueId,
                        folder.DisplayName,
                        folderPath,
                        folder.TotalCount,
                        folder.UnreadCount
                    ));

                    // Recursively get child folders
                    if (folder.ChildFolderCount > 0)
                    {
                        await GetFoldersRecursiveAsync(service, folder.Id, folderPath, folders);
                    }
                }

                folderView.Offset = findResults.NextPageOffset ?? 0;

            } while (findResults.MoreAvailable);
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Warning: Error retrieving child folders: {ex.Message}");
        }
    }

    /// <summary>
    /// Exports emails from an archive folder
    /// </summary>
    public async System.Threading.Tasks.Task<List<EmailItem>> ExportArchiveEmailsAsync(
        string mailboxEmail,
        string folderId,
        int count)
    {
        Console.WriteLine($"\n[EWS] Exporting emails from archive...");
        Console.WriteLine($"  Mailbox: {mailboxEmail}");
        Console.WriteLine($"  Count: {(count == 0 ? "all" : count.ToString())}");

        var (service, _) = await CreateEwsServiceAsync();
        var emails = new List<EmailItem>();

        try
        {
            var folder = await SystemTask.Run(() => Folder.Bind(service, new FolderId(folderId)));

            Console.WriteLine($"✓ Folder bound: {folder.DisplayName}");

            // Determine how many emails to retrieve
            int pageSize = count == 0 ? 1000 : Math.Min(count, 1000);
            var itemView = new ItemView(pageSize)
            {
                PropertySet = new PropertySet(BasePropertySet.FirstClassProperties, ItemSchema.Body)
            };

            int retrieved = 0;
            bool exportAll = (count == 0);

            do
            {
                var findResults = await SystemTask.Run(() => service.FindItems(folder.Id, itemView));

                foreach (var item in findResults.Items.OfType<EmailMessage>())
                {
                    emails.Add(new EmailItem
                    {
                        Id = item.Id.UniqueId,
                        Subject = item.Subject,
                        From = new EmailAddress
                        {
                            Name = item.From?.Name,
                            Address = item.From?.Address
                        },
                        ToRecipients = item.ToRecipients?.Select(r => new EmailAddress
                        {
                            Name = r.Name,
                            Address = r.Address
                        }).ToList(),
                        CcRecipients = item.CcRecipients?.Select(r => new EmailAddress
                        {
                            Name = r.Name,
                            Address = r.Address
                        }).ToList(),
                        BccRecipients = item.BccRecipients?.Select(r => new EmailAddress
                        {
                            Name = r.Name,
                            Address = r.Address
                        }).ToList(),
                        ReceivedDateTime = item.DateTimeReceived,
                        SentDateTime = item.DateTimeSent,
                        HasAttachments = item.HasAttachments,
                        Importance = item.Importance.ToString(),
                        IsRead = item.IsRead,
                        IsDraft = item.IsDraft,
                        InternetMessageId = item.InternetMessageId,
                        ConversationId = item.ConversationId?.UniqueId,
                        Categories = item.Categories?.ToList(),
                        Body = new EmailBody
                        {
                            ContentType = item.Body.BodyType == BodyType.HTML ? "Html" : "Text",
                            Content = item.Body.Text
                        },
                        BodyPreview = item.Preview
                    });

                    retrieved++;

                    if (!exportAll && retrieved >= count)
                    {
                        break;
                    }
                }

                if (exportAll && findResults.MoreAvailable)
                {
                    itemView.Offset = findResults.NextPageOffset ?? 0;

                    if (retrieved % 1000 == 0)
                    {
                        Console.WriteLine($"Retrieved {retrieved} emails...");
                    }
                }
                else
                {
                    break;
                }

            } while (exportAll);

            Console.WriteLine($"✓ Exported {emails.Count} email(s) from archive");

            return emails;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"✗ Error exporting archive emails: {ex.Message}");
            throw;
        }
    }
}

/// <summary>
/// Email item model (compatible with Graph API model structure)
/// </summary>
public class EmailItem
{
    public string? Id { get; set; }
    public string? Subject { get; set; }
    public EmailAddress? From { get; set; }
    public List<EmailAddress>? ToRecipients { get; set; }
    public List<EmailAddress>? CcRecipients { get; set; }
    public List<EmailAddress>? BccRecipients { get; set; }
    public DateTime? ReceivedDateTime { get; set; }
    public DateTime? SentDateTime { get; set; }
    public bool HasAttachments { get; set; }
    public string? Importance { get; set; }
    public bool IsRead { get; set; }
    public bool IsDraft { get; set; }
    public string? InternetMessageId { get; set; }
    public string? ConversationId { get; set; }
    public List<string>? Categories { get; set; }
    public EmailBody? Body { get; set; }
    public string? BodyPreview { get; set; }
}

public class EmailAddress
{
    public string? Name { get; set; }
    public string? Address { get; set; }
}

public class EmailBody
{
    public string? ContentType { get; set; }
    public string? Content { get; set; }
}
