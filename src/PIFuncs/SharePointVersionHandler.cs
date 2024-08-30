using System;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using PnP.Core.Services;
using PnP.Core.Model.SharePoint;
using Microsoft.Extensions.Configuration;
using PnP.Core.QueryModel;

namespace PIFunc
{
    public class SharePointFileHandler
    {
        private readonly IPnPContextFactory _pnpContextFactory;
        private readonly IConfiguration _configuration;

        public SharePointFileHandler(IPnPContextFactory pnpContextFactory, IConfiguration configuration)
        {
            _pnpContextFactory = pnpContextFactory;
            _configuration = configuration;
        }

        public async Task MoveLowerVersionsToArchiveAsync(string listTitle, string fileName, string docid)
        {
            try
            {
                string siteUrl = _configuration["SiteUrl"];
                string archiveLibraryName = _configuration["ArchiveLibraryName"] ?? "ArchiveDocument";
                string publishedDocumentsFolder = listTitle ?? "PublishedDocument"; 

                if (string.IsNullOrWhiteSpace(listTitle) || string.IsNullOrWhiteSpace(fileName))
                {
                    Console.WriteLine("List title and file name are required.");
                    return;
                }

                using (var context = await _pnpContextFactory.CreateAsync(new Uri(siteUrl)))
                {
                    // Get the list by title
                    var library = await context.Web.Lists.GetByTitleAsync(listTitle);

                    // Retrieve all items from the list
                    var items = await library.Items.ToListAsync();

                    // Filter items by docid
                    var relevantItems = items
                        .Where(item => item.Values.ContainsKey("docid") && item.Values["docid"]?.ToString() == docid)
                        .Select(item => new
                        {
                            Item = item,
                            FileName = item.Values.ContainsKey("Title") ? item.Values["Title"].ToString() : null,
                            // Construct the ServerRelativeUrl
                            ServerRelativeUrl = ConstructServerRelativeUrl(siteUrl, publishedDocumentsFolder, item.Values.ContainsKey("Title") ? item.Values["Title"].ToString() : "default_filename"),

                            RevisionNo = item.Values.ContainsKey("RevisionNo") ? item.Values["RevisionNo"].ToString() : null
                        })
                        .OrderBy(item => item.RevisionNo)
                        .ToList();

                    if (!relevantItems.Any())
                    {
                        Console.WriteLine("No items found for the specified docid.");
                        return;
                    }

                    // Determine the latest revision number
                    var latestRevisionNo = relevantItems.LastOrDefault()?.RevisionNo;

                    // Move items with lower revision numbers to the archive
                    foreach (var revItem in relevantItems)
                    {
                        if (revItem.RevisionNo != latestRevisionNo && revItem.ServerRelativeUrl != null)
                        {
                            string sourceUrl = revItem.ServerRelativeUrl; 
                            string destinationUrl = $"/sites/testPortal/{archiveLibraryName}/{Path.GetFileName(sourceUrl)}";

                            await MoveFileWithMetadataAsync(context, revItem.Item, sourceUrl, destinationUrl, archiveLibraryName);
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while moving lower versions to archive: {ex.Message}");
            }
        }

        private string ConstructServerRelativeUrl(string siteUrl, string folderName, string fileName)
        {
            // Ensure siteUrl ends with a slash
            if (!siteUrl.EndsWith("/"))
            {
                siteUrl += "/";
            }

            // Extract the site-relative part from the siteUrl
            var uri = new Uri(siteUrl);
            var siteRelativePath = uri.AbsolutePath.TrimEnd('/'); // Remove any trailing slash

            // Construct ServerRelativeUrl
            var serverRelativeUrl = $"{siteRelativePath}/{folderName}/{fileName}".Replace("//", "/"); 

            return serverRelativeUrl;
        }

        private async Task MoveFileWithMetadataAsync(PnPContext context, IListItem sourceItem, string sourceUrl, string destinationUrl, string archiveLibraryName)
        {
            try
            {
                Console.WriteLine($"Attempting to get file from URL: {sourceUrl}");

                if (string.IsNullOrWhiteSpace(sourceUrl))
                {
                    throw new ArgumentException("Source URL is null or empty.");
                }

                // Attempt to get the file from SharePoint using ServerRelativeUrl
                var file = await context.Web.GetFileByServerRelativeUrlAsync(sourceUrl);
                if (file == null)
                {
                    throw new Exception("File not found at the specified URL.");
                }

                Console.WriteLine($"Copying file to destination URL: {destinationUrl}");

                // Copy the file to the destination
                await file.CopyToAsync(destinationUrl, overwrite: true);

                // Retrieve the archive library
                var archiveLibrary = await context.Web.Lists.GetByTitleAsync(archiveLibraryName);

                // Get the copied file
                var copiedFile = await context.Web.GetFileByServerRelativeUrlAsync(destinationUrl);

                // Explicitly load the necessary properties
                await copiedFile.LoadAsync(f => f.ListItemAllFields);

                // Ensure that the ListItemAllFields property is loaded correctly
                var copiedItem = await copiedFile.ListItemAllFields.GetAsync();

                Console.WriteLine($"Retrieved ListItemAllFields: {copiedItem.Id}"); // Debugging output

                // Update the copied item with source item fields
                foreach (var field in sourceItem.Values)
                {
                    // Skip read-only fields
                    if (!field.Key.Equals("FileRef", StringComparison.OrdinalIgnoreCase) &&
                        !field.Key.Equals("ID", StringComparison.OrdinalIgnoreCase) &&
                        !field.Key.Equals("Created", StringComparison.OrdinalIgnoreCase) &&
                        !field.Key.Equals("Modified", StringComparison.OrdinalIgnoreCase))
                    {
                        try
                        {
                            // Check if the property exists before setting it
                            if (copiedItem.Values.ContainsKey(field.Key))
                            {
                                copiedItem[field.Key] = field.Value;
                            }
                            else
                            {
                                Console.WriteLine($"Property {field.Key} does not exist on the copied item.");
                            }
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine($"Error updating property {field.Key}: {ex.Message}");
                        }
                    }
                }

                // Update the copied item
                await copiedItem.UpdateAsync();

                await file.DeleteAsync();

                Console.WriteLine("File moved successfully.");
            }
            catch (PnP.Core.ClientException ex)
            {
                Console.WriteLine($"A PnP Core ClientException occurred: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
            catch (Exception ex)
            {
                Console.WriteLine($"An error occurred while moving the file: {ex.Message}");
                Console.WriteLine($"Stack Trace: {ex.StackTrace}");
            }
        }







    }
}
