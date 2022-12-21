using System.IO;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Azure.WebJobs;
using Microsoft.Azure.WebJobs.Extensions.Http;
using Microsoft.Extensions.Logging;
using PnP.Core.Services;
using System.Net.Http;

using System.Linq;
using PnP.Core.Model.SharePoint;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using System;
using PnP.Core.QueryModel;
using System.Collections.Concurrent;


namespace Demo.Function
{
    using System.Collections.Generic;
    using DocumentFormat.OpenXml.Wordprocessing;

    public class HttpTriggerSharepointServices
    {
        private const string APPROVAL_HISTORY_LIST_NAME = "ApprovalHistory";
        private const int BORDER_WIDTH = 1;
        readonly IPnPContextFactory _pnpContextFactory;
        private static ConcurrentDictionary<string, string> _runningTasks = new ConcurrentDictionary<string, string>();
        public HttpTriggerSharepointServices(IPnPContextFactory pnpContextFactory, ILogger<HttpTriggerSharepointServices> logger)
        {
            _pnpContextFactory = pnpContextFactory;
            Logger = logger;
        }

        public ILogger<HttpTriggerSharepointServices> Logger { get; }

        /// <summary>
        /// Ping request for alive status
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        [FunctionName("HttpTriggerPing")]
        public async Task<IActionResult> RunPing(
           [HttpTrigger(AuthorizationLevel.Anonymous, "get", "post", Route = null)] HttpRequestMessage req)
        {
            await Task.CompletedTask;
            return new JsonResult(new { Status = "Running" });
        }

        /// <summary>
        /// Adds approval history of a office 365 docx from a sharepoint library, run against requested site : TestPortal, configured SiteUrl
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>

        [FunctionName("HttpTrigger1MyFunc2")]
        public async Task<IActionResult> Run2(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req)
        {

            #region Parameters
            Logger.LogInformation("C# HTTP trigger function processed a request.");

            try
            {
                var result = new { };
                string docid, libname, flname, destSpFolder, targetSite;
                bool download;

                ReadDocumentApprovalHistoryParameters(req, out docid, out libname, out flname, out download, out destSpFolder, out targetSite);

                #endregion

                using (var ctx = await _pnpContextFactory.CreateAsync(targetSite))
                {
                    var fileInfo = await QueryFileAndMetaData(libname, flname, ctx);

                    if (string.IsNullOrWhiteSpace(docid))
                        docid = fileInfo.Id.ToString();

                    var destinationLibrary = await ctx.Web.Lists.GetByTitleAsync(destSpFolder, l => l.RootFolder);

                    var shareDocuments = await ctx.Web.Lists.GetByTitleAsync(libname, l => l.RootFolder);

                    IFile docx = shareDocuments.RootFolder.Files.FirstOrDefault(o => o.Name == flname);

                    var folderContents = await shareDocuments.RootFolder.GetAsync(o => o.Files);

                    if (docx == null || string.IsNullOrEmpty(docid))
                        return new NotFoundResult();

                    else
                    {
                        IEnumerable<IListItem> historyItems = await GetApprovalHistory(docid, ctx, fileInfo.RevisionNo);

                        var bytes = docx.GetContentBytes();
                        var tmpflName = Guid.NewGuid().ToString();
                        var tmpDocx = Path.Combine(Path.GetTempPath(), $"{tmpflName}.docx");
                        File.WriteAllBytes(tmpDocx, bytes);
                        using (var doc = WordprocessingDocument.Open(tmpDocx, true))
                        {
                            Table table;
                            OpenXmlAttribute attrib;

                            CleanExistingTable(doc, out table, out attrib);

                            table = CreateApprovalHistoryTable(attrib);
                            AppendApprovalHistory(historyItems, doc, table);
                        }

                        try
                        {
                            await PublishDocument(flname, destinationLibrary, tmpDocx);
                        }
                        catch (Exception ex)
                        {
                            Console.Error.WriteLine(ex.Message);
                        }

                        if (download)
                        {
                            return DownloadPublishedDocument(flname, tmpDocx);
                        }
                        else
                            return new JsonResult(new { Success = true });
                    }


                }
            }
            catch (Exception ex)
            {
                Logger.LogCritical(ex, ex.Message);
                throw;
            }
        }

        private static void ReadDocumentApprovalHistoryParameters(HttpRequestMessage req, out string docid, out string libname, out string flname, out bool download, out string destSpFolder, out string targetSite)
        {
            docid = string.Empty;
            libname = "DmsDocument";
            flname = "CM-UDR-1-V2.docx";
            download = false;
            var debug = false;
            destSpFolder = "PublishedDocument";
            var qry = req.RequestUri.ParseQueryString().GetValues("d");
            var qryd = req.RequestUri.ParseQueryString().GetValues("dwnld");
            var qrylib = req.RequestUri.ParseQueryString().GetValues("lib");
            var qDocId = req.RequestUri.ParseQueryString().GetValues("docid");
            var qDebug = req.RequestUri.ParseQueryString().GetValues("debug");
            var qPFolder = req.RequestUri.ParseQueryString().GetValues("pfolder");

            if (qry != null)
                flname = qry.FirstOrDefault();

            if (qryd != null)
                bool.TryParse(qryd.FirstOrDefault(), out download);

            if (qrylib != null)
                libname = qrylib.FirstOrDefault();

            if (qDocId != null)
                docid = qDocId.FirstOrDefault();

            if (qDebug != null)
                bool.TryParse(qDebug.FirstOrDefault(), out debug);

            if (qPFolder != null)
                destSpFolder = qPFolder.FirstOrDefault();

            targetSite = "Default";
            if (debug)
                targetSite = "TestPortal";
        }

        private static IActionResult DownloadPublishedDocument(string flname, string tmpDocx)
        {
            return new PhysicalFileResult(tmpDocx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
            {
                FileDownloadName = flname
            };
        }

        private async Task PublishDocument(string flname, IList destinationLibrary, string tmpDocx)
        {
            using (Stream s = new FileStream(tmpDocx, FileMode.Open))
            {
                try
                {
                    // save to destination folder
                    await destinationLibrary.RootFolder.Files.AddAsync(flname, s, true);
                }
                catch (System.Exception ex)
                {
                    Logger.LogError(ex, ex.Message + ex.StackTrace);

                }
            }
        }

        private static void AppendApprovalHistory(IEnumerable<IListItem> historyItems, WordprocessingDocument doc, Table table)
        {
            foreach (var item in historyItems)
            {
                var level = Convert.ToString(item["Level"]);
                var role = Convert.ToString(item["Role"]);
                var approvalDate = Convert.ToDateTime(item["Created"]).ToString("dd-MMM-yyyy");
                var approve = Convert.ToString(item["UserName"]);

                TableRow tr = new TableRow();
                TableRowProperties trProp = new TableRowProperties(new TableRowHeight
                {
                    HeightType = new EnumValue<HeightRuleValues>(HeightRuleValues.Auto),
                });

                CreateCell(level, tr);
                CreateCell(role, tr);
                CreateCell(approve, tr, false, 2160);
                CreateCell(approvalDate, tr);

                table.Append(tr);

            }

            // Append the table to the document.
            doc.MainDocumentPart.Document.Body.Append(table);
            doc.Save();
        }

        private static Table CreateApprovalHistoryTable(OpenXmlAttribute attrib)
        {
            Table table = new Table();
            table.SetAttribute(attrib);

            TableProperties tblProp = CreateTableProperties();
            table.AppendChild<TableProperties>(tblProp);

            TableRow trHead = new TableRow();

            CreateCell("Level in Route", trHead, true);
            CreateCell("Role/Designation", trHead, true);
            CreateCell("Name of the Approver", trHead, true);
            CreateCell("Date of Approval", trHead, true);

            table.Append(trHead);
            return table;
        }

        private static void CleanExistingTable(WordprocessingDocument doc, out Table table, out OpenXmlAttribute attrib)
        {
            var tables = doc.MainDocumentPart.Document.Body.Elements<Table>();

            table = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault(o => o.LocalName == "tbl");
            attrib = new OpenXmlAttribute("tbl", "history", "", "table");
            if (table != null)
            {
                doc.MainDocumentPart.Document.Body.RemoveChild(table);
                doc.Save();
            }
        }

        private async Task<IEnumerable<IListItem>> GetApprovalHistory(string docid, PnPContext ctx, string revision)
        {
            string viewXml = @"<View>
                    <Query>
                      <Where>
                        <And>
                        <Eq>
                          <FieldRef Name='DMSID'/>
                          <Value Type='text'>" + docid + @"</Value>
                        </Eq>
                        <Eq>
                          <FieldRef Name='RevisionNo'/>
                          <Value Type='text'>" + revision + @"</Value>
                        </Eq>
                       </And>
                      </Where>
                    </Query>
                   </View>";

            var approvalhistory = ctx.Web.Lists.GetByTitle(APPROVAL_HISTORY_LIST_NAME);
            await approvalhistory.LoadItemsByCamlQueryAsync(new CamlQueryOptions
            {
                ViewXml = viewXml,
                DatesInUtc = false,
            });

            var historyItems = approvalhistory.Items.AsRequested();

            if (historyItems?.Count() == 0)
            {
                Logger.LogError($"Approval History for the document {docid} not found");
                throw new Exception("Approval history not found");
            }

            return historyItems;
        }

        private async Task<(string DocumentName, string RevisionNo, int Id, string DocID)> QueryFileAndMetaData(string libname, string flname, PnPContext ctx)
        {
            // Assume the fields where not yet loaded, so loading them with the list
            var myList = ctx.Web.Lists.GetByTitle(libname, p => p.Title,
                                                                 p => p.Fields.QueryProperties(p => p.InternalName,
                                                                                               p => p.FieldTypeKind,
                                                                                               p => p.TypeAsString,
                                                                                               p => p.Title));

            // Build a query that only returns the Title field for items where the Title field starts with "Item1"
            string viewXml1 = @"<View>
                    <ViewFields>
                      <FieldRef Name='Title' />
                      <FieldRef Name='Name' />
                      <FieldRef Name='FileRef' />
                      <FieldRef Name='RevisionNo' />
                      <FieldRef Name='DocID' />
                    </ViewFields>
                    <Query>
                        <Where>
                        <Eq>
                          <FieldRef Name='Title'/>
                          <Value Type='text'>" + flname + @"</Value>
                        </Eq>
                      </Where>
                    </Query>
                    <OrderBy Override='TRUE'><FieldRef Name= 'ID' Ascending= 'FALSE' /></OrderBy>
                   </View>";

            await myList.LoadItemsByCamlQueryAsync(new CamlQueryOptions()
            {
                ViewXml = viewXml1,
                DatesInUtc = true
            }, p => p.FieldValuesAsText, p => p.RoleAssignments.QueryProperties(p => p.PrincipalId, p => p.RoleDefinitions));


            var items = myList.Items.AsRequested();

            if (items.Any())
            {
                var doc = items.FirstOrDefault();

                return (doc.Title, doc.FieldValuesAsText["RevisionNo"]?.ToString(), doc.Id, doc.FieldValuesAsText["DocID"]?.ToString());
            }

            return default((string, string, int, string));
        }

        /// <summary>
        /// Adds approval history of a office 365 docx from a sharepoint library, https://credentinfotec.sharepoint.com/sites/demo/ihub
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>

        [FunctionName("HttpTrigger1MyFunc3")]
        public async Task<IActionResult> Run3(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req)
        {
            Logger.LogInformation("C# HTTP trigger function processed a request.");

            var result = new { };

            var libname = "Documents";
            var flname = "Document.docx";
            var download = false;
            var qry = req.RequestUri.ParseQueryString().GetValues("d");
            var qryd = req.RequestUri.ParseQueryString().GetValues("dwnld");
            var qrylib = req.RequestUri.ParseQueryString().GetValues("lib");
            if (qry != null)
                flname = qry.FirstOrDefault();

            if (qryd != null)
                bool.TryParse(qryd.FirstOrDefault(), out download);

            if (qrylib != null)
                libname = qrylib.FirstOrDefault();

            using (var ctx = await _pnpContextFactory.CreateAsync("Default"))
            {

                var approvalhistory = ctx.Web.Lists.GetByTitle(APPROVAL_HISTORY_LIST_NAME);
                var historyItems = approvalhistory.Items.Where(o => o.Title == flname).ToList();

                var shareDocuments = await ctx.Web.Lists.GetByTitleAsync(libname, l => l.RootFolder);
                var folderContents = await shareDocuments.RootFolder.GetAsync(o => o.Files);

                var documents = from d in folderContents.Files.AsEnumerable()
                                select new
                                {
                                    d.Name,
                                    d.TimeLastModified,
                                    d.UniqueId
                                };

                IFile docx = null;
                var tmpLocation = System.IO.Path.GetTempPath();

                foreach (var fl in folderContents.Files.AsEnumerable())
                {
                    if (fl.Name == flname)
                    {
                        Logger.LogInformation("Docx found");
                        Logger.LogInformation(fl.Name);
                        docx = fl;
                        // return new NotFoundResult();
                    }
                }

                if (docx == null)
                    return new NotFoundResult();

                else
                {
                    var bytes = docx.GetContentBytes();
                    var tmpflName = Guid.NewGuid().ToString();
                    var tmpDocx = Path.Combine(Path.GetTempPath(), $"{tmpflName}.docx");
                    File.WriteAllBytes(tmpDocx, bytes);
                    using (var doc = WordprocessingDocument.Open(tmpDocx, true))
                    {
                        var tables = doc.MainDocumentPart.Document.Body.Elements<Table>();

                        var table = doc.MainDocumentPart.Document.Body.Elements<Table>().FirstOrDefault(o => o.LocalName == "tbl");

                        var attrib = new OpenXmlAttribute("tbl", "history", "", "table");

                        if (table != null)
                        {
                            doc.MainDocumentPart.Document.Body.RemoveChild(table);
                            doc.Save();
                        }

                        table = new Table();


                        TableProperties tblProp = CreateTableProperties();

                        // Append the TableProperties object to the empty table.
                        table.AppendChild<TableProperties>(tblProp);

                        table.SetAttribute(attrib);

                        TableRow trHead = new TableRow();


                        CreateCell("Role", trHead);
                        CreateCell("Name", trHead);
                        CreateCell("DateTime", trHead);
                        table.Append(trHead);

                        foreach (var item in approvalhistory.Items)
                        {
                            var val1 = Convert.ToString(item["Title"]);
                            var val2 = Convert.ToString(item["Role"]);
                            var val3 = Convert.ToString(item["Name"]);
                            var val4 = Convert.ToString(item["DateTime"]);
                            TableRow tr = new TableRow();


                            CreateCell(val2, tr);
                            CreateCell(val3, tr);
                            CreateCell(val4, tr);
                            table.Append(tr);

                        }


                        // Append the table to the document.
                        doc.MainDocumentPart.Document.Body.Append(table);

                        doc.Save();
                    }

                    try
                    {
                        using (Stream s = new FileStream(tmpDocx, FileMode.Open))
                        {
                            try
                            {
                                await folderContents.Files.AddAsync(flname, s, true);
                            }
                            catch (System.Exception ex)
                            {
                                Logger.LogError(ex, ex.Message + ex.StackTrace);

                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.Error.WriteLine(ex.Message);
                        //throw;
                    }

                    if (download)
                    {
                        return new PhysicalFileResult(tmpDocx, "application/vnd.openxmlformats-officedocument.wordprocessingml.document")
                        {
                            FileDownloadName = flname
                        };
                    }
                    else
                        return new JsonResult(new { Success = true });
                }


            }
        }

        private static TableProperties CreateTableProperties()
        {

            // Create a TableProperties object and specify its border information.
            return new TableProperties(new TableWidth()
            {
                Type = new EnumValue<TableWidthUnitValues>(TableWidthUnitValues.Pct),
                Width = "100%"
            },
                new TableBorders(
                    new TopBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = BORDER_WIDTH
                    },
                    new BottomBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = BORDER_WIDTH
                    },
                    new LeftBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = BORDER_WIDTH
                    },
                    new RightBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = BORDER_WIDTH
                    },
                    new InsideHorizontalBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = BORDER_WIDTH
                    },
                    new InsideVerticalBorder()
                    {
                        Val =
                        new EnumValue<BorderValues>(BorderValues.Single),
                        Size = BORDER_WIDTH
                    }
                )
            );
        }

        private static void CreateCell(string val1, TableRow tr, bool boldText = false, uint width = 1440)
        {
            TableCell tc = new TableCell();

            // Specify the width property of the table cell.
            if (width == 0)
                tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Nil }));
            else
                tc.Append(new TableCellProperties(
                new TableCellWidth() { Type = TableWidthUnitValues.Dxa, Width = width.ToString() }));

            var text = new Text(val1);
            var run = new Run();
            var runProp = new RunProperties();

            if (boldText)
                runProp.Append(new Bold());

            run.Append(runProp);
            run.Append(text);
            // Specify the table cell content.
            tc.Append(new Paragraph(run));

            // Append the table cell to the table row.
            tr.Append(tc);
        }
    }

}
