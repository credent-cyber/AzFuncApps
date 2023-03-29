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
using System.Collections.Generic;

namespace Demo.Function
{

    using DocumentFormat.OpenXml.Wordprocessing;
    using PIFunc.DocxHelper;
    using XlsxHelper;

    public class HttpTriggerSharepointServices
    {
        private const string APPROVAL_HISTORY_LIST_NAME = "ApprovalHistory";
        private const int BORDER_WIDTH = 1;
        readonly IPnPContextFactory _pnpContextFactory;
        private static ConcurrentDictionary<string, string> _runningTasks = new ConcurrentDictionary<string, string>();
        public HttpTriggerSharepointServices(IPnPContextFactory pnpContextFactory,
            ILogger<HttpTriggerSharepointServices> logger, AzureFunctionSettings functionSettings)
        {
            _pnpContextFactory = pnpContextFactory;
            Logger = logger;
            FunctionSettings = functionSettings;
        }

        public ILogger<HttpTriggerSharepointServices> Logger { get; }
        public AzureFunctionSettings FunctionSettings { get; }

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
        /// Adds approval history of a office 365 docx, xlsx from a sharepoint library, run against requested site : TestPortal, configured SiteUrl
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
                        var isXlsx = Path.GetExtension(flname) == ".xlsx" ? true : false;
                        var isDocx = Path.GetExtension(flname) == ".docx" ? true : false;

                        IEnumerable<IListItem> historyItems = await GetApprovalHistory(docid, ctx, fileInfo.RevisionNo);

                        var bytes = docx.GetContentBytes();
                        var tmpflName = Guid.NewGuid().ToString();
                        var ext = isXlsx ? ".xlsx" : ".docx";
                        var tmpDocx = Path.Combine(Path.GetTempPath(), $"{tmpflName}.{ext}");

                        File.WriteAllBytes(tmpDocx, bytes);

                        if (isDocx)
                        {
                            using (var doc = WordprocessingDocument.Open(tmpDocx, true))
                            {
                                Table table;

                                OpenXmlAttribute attrib;

                                CleanExistingTable(doc, out table, out attrib);

                                table = CreateApprovalHistoryTable(attrib);
                                AppendApprovalHistory(historyItems, doc, table);

                                var versionDate = GetVersionDate(historyItems);
                                var table2 = CreateMetaDataTable(attrib, docid, fileInfo.ProcedureRef, fileInfo.RevisionNo.ToUpper(), versionDate, fileInfo.FileName, string.Empty);

                                DocumentHeader.AddMetadata(doc, table2);
                            }

                        }

                        if (isXlsx)
                        {
                            var auditHistory = new AuditHistory();
                            var data  = GetHistoryDataArray(historyItems);
                            auditHistory.Append(tmpDocx, null, data);
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


        /// <summary>
        /// Wrapper function for Run2 - new name with same functionality
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        [FunctionName("HttpTriggerDocxApprovalHistory")]
        public async Task<IActionResult> Run3(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req)
        {
            return await Run2(req);
        }

        /// <summary>
        /// Wrapper function for Run2 - new name with same functionality
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        [FunctionName("HttpTriggerXlsxApprovalHistory")]
        public async Task<IActionResult> Run4(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req)
        {
            return await Run2(req);
        }


        /// <summary>
        /// CreateTableProperties
        /// </summary>
        /// <returns></returns>
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

        private static void CreateCell(string val1, TableRow tr, bool boldText = false, uint width = 1440, TableCellProperties tableCellProperties = null)
        {
            TableCell tc = new TableCell();

            if(tableCellProperties != null)
                tc.Append(tableCellProperties);
            
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


        /// <summary>
        /// ReadDocumentApprovalHistoryParameters
        /// </summary>
        /// <param name="req"></param>
        /// <param name="docid"></param>
        /// <param name="libname"></param>
        /// <param name="flname"></param>
        /// <param name="download"></param>
        /// <param name="destSpFolder"></param>
        /// <param name="targetSite"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="flname"></param>
        /// <param name="destinationLibrary"></param>
        /// <param name="tmpDocx"></param>
        /// <returns></returns>
        private async Task PublishDocument(string flname, PnP.Core.Model.SharePoint.IList destinationLibrary, string tmpDocx)
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

        private List<string[]> GetHistoryDataArray(IEnumerable<IListItem> historyItems)
        {
            List<string[]> historyDataArray = new List<string[]>();

            foreach(var item in historyItems)
            {
                var itemArray = new List<string>();
                var level = Convert.ToString(item["Level"]);
                var role = Convert.ToString(item["Role"]);
                var action = Convert.ToString(item["Action"]);
                var approvalDate = Convert.ToDateTime(item["Created"]).ToString("dd-MMM-yyyy");
                var approver = Convert.ToString(item["UserName"]);
                var designation = Convert.ToString(item["DMSRole"]);

                var isFiltered = FunctionSettings
                    .ApprovalHistoryExcludedRole.Any(o => o.Equals(role, StringComparison.InvariantCultureIgnoreCase));

                isFiltered = isFiltered || FunctionSettings
                    .ApprovalHistoryExcludedAction.Any(o => o.Equals(action, StringComparison.InvariantCultureIgnoreCase));

                if (isFiltered)
                    continue;
               
                role = $"{role}/{designation}";
                
                itemArray.Add(level);
                itemArray.Add(role);
                itemArray.Add(approver);
                itemArray.Add(approvalDate);

                historyDataArray.Add(itemArray.ToArray());

            }

            historyDataArray = historyDataArray.OrderBy(arr => arr[0])
                .ToList();

            return historyDataArray;

        }

        private string GetVersionDate(IEnumerable<IListItem> historyItems)
        {
            var item = historyItems.FirstOrDefault();

            return Convert.ToDateTime(item["Created"]).ToString("dd-MMM-yyyy");

        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="historyItems"></param>
        /// <param name="doc"></param>
        /// <param name="table"></param>
        private void AppendApprovalHistory(IEnumerable<IListItem> historyItems, WordprocessingDocument doc, Table table)
        {
            foreach (var item in historyItems)
            {
                var level = Convert.ToString(item["Level"]);
                var role = Convert.ToString(item["Role"]);
                var action = Convert.ToString(item["Action"]);
                var approvalDate = Convert.ToDateTime(item["Created"]).ToString("dd-MMM-yyyy");
                var approve = Convert.ToString(item["UserName"]);
                var designation = Convert.ToString(item["DMSRole"]);

                if (FunctionSettings.ApprovalHistoryExcludedRole.Any(o => o.Equals(role, StringComparison.InvariantCultureIgnoreCase)))
                {
                    Logger.LogWarning($"Role: [{role}] is filtered from the configuration settings");
                    continue;
                }

                if (FunctionSettings.ApprovalHistoryExcludedAction.Any(o => o.Equals(action, StringComparison.InvariantCultureIgnoreCase)))
                {
                    Logger.LogWarning($"Action: [{action}] is filtered from the configuration settings");
                    continue;
                }

                TableRow tr = new TableRow();
                TableRowProperties trProp = new TableRowProperties(new TableRowHeight
                {
                    HeightType = new EnumValue<HeightRuleValues>(HeightRuleValues.Auto),
                });

                CreateCell(level, tr);

                role = $"{role}/{designation}";

                CreateCell(role, tr);
                CreateCell(approve, tr, false, 2160);
                CreateCell(approvalDate, tr);

                table.Append(tr);

            }

            // Append the table to the document.
            doc.MainDocumentPart.Document.Body.Append(table);
            
            doc.Save();
        }


        /// <summary>
        /// 
        /// </summary>
        /// <param name="attrib"></param>
        /// <returns></returns>
        private static Table CreateApprovalHistoryTable(OpenXmlAttribute attrib)
        {
            Table table = new Table();
            table.SetAttribute(attrib);

            TableProperties tblProp = CreateTableProperties();
            table.AppendChild(tblProp);

            TableRow trHead = new TableRow();

            CreateCell("Level in Route", trHead, true);
            CreateCell("Role/Designation", trHead, true);
            CreateCell("Name of the Approver", trHead, true);
            CreateCell("Date of Approval", trHead, true);

            table.Append(trHead);
            return table;
        }

        private static Table CreateMetaDataTable(OpenXmlAttribute attrib, string docId, 
            string procedureReference, string version, string revisionDate, string content, string copyNumber)
        {
            Table table = new Table();
            table.SetAttribute(attrib);

            TableProperties tblProp = CreateTableProperties();
            table.AppendChild(tblProp);

            TableRow trHead = new TableRow();

            CreateCell("DOC ID NO:", trHead, true);
            CreateCell(docId, trHead, false);
            CreateCell("PROCEDURE REF NO:", trHead, true);
            CreateCell(procedureReference, trHead, false);

            table.Append(trHead);


            TableRow row2 = new TableRow();

            CreateCell("REVISION NO:", row2, true);
            CreateCell(version, row2, false);
            CreateCell("REVISION DATE:", row2, true);
            CreateCell(revisionDate, row2, false);

            table.Append(row2);

            TableRow row3 = new TableRow();

            TableCellProperties cellOneProperties = new TableCellProperties();
            cellOneProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Restart
            });

            TableCellProperties cellTwoProperties = new TableCellProperties();
            cellTwoProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Continue
            });

            TableCellProperties cellThreeProperties = new TableCellProperties();
            cellThreeProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Continue
            });

            TableCellProperties cellFourProperties = new TableCellProperties();
            cellFourProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Continue
            });

            TableCellProperties cellFiveProperties = new TableCellProperties();
            cellFiveProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Restart
            });

            TableCellProperties cellSixProperties = new TableCellProperties();
            cellSixProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Continue
            });

            TableCellProperties cellSevenProperties = new TableCellProperties();
            cellSevenProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Restart
            });

            TableCellProperties cellEightProperties = new TableCellProperties();
            cellEightProperties.Append(new HorizontalMerge()
            {
                Val = MergedCellValues.Continue
            });


            CreateCell(content, row3, true, tableCellProperties: cellOneProperties);
            CreateCell("", row3, false, tableCellProperties: cellTwoProperties);
            CreateCell("", row3, false, tableCellProperties: cellThreeProperties);
            CreateCell("", row3, false, tableCellProperties: cellFourProperties);

            table.Append(row3);

            TableRow row4 = new TableRow();

            CreateCell("Controlled if stamped in red", row4, true, tableCellProperties: cellFiveProperties);
            CreateCell("", row4, true, tableCellProperties: cellSixProperties);

            CreateCell("COPY NO.", row4, true, tableCellProperties: cellSevenProperties);
            CreateCell(copyNumber, row4, true, tableCellProperties: cellEightProperties);

            table.Append(row4);
            return table;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="table"></param>
        /// <param name="attrib"></param>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="docid"></param>
        /// <param name="ctx"></param>
        /// <param name="revision"></param>
        /// <returns></returns>
        /// <exception cref="Exception"></exception>
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
                     <OrderBy>
                          <FieldRef Name='Created' Ascending='False' />
                     </OrderBy>
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

        /// <summary>
        /// 
        /// </summary>
        /// <param name="libname"></param>
        /// <param name="flname"></param>
        /// <param name="ctx"></param>
        /// <returns></returns>
        private async Task<(string DocumentName, string RevisionNo, int Id, string DocID, string FileName, string ProcedureRef)> QueryFileAndMetaData(string libname, string flname, PnPContext ctx)
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
                      <FieldRef Name='DocumentName' />
                      <FieldRef Name='ProcedureRef' />
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

                return (doc.Title, doc.FieldValuesAsText["RevisionNo"]?.ToString(), doc.Id, doc.FieldValuesAsText["DocID"]?.ToString(), doc.FieldValuesAsText["DocumentName"]?.ToString(), doc.FieldValuesAsText["ProcedureRef"]?.ToString());
            }

            return default((string, string, int, string, string, string));
        }
    }

}
