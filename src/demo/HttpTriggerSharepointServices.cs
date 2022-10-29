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
using Demo.ITextSharp;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;

using DocumentFormat.OpenXml.Wordprocessing;
using System;
using PnP.Core.QueryModel;
using System.Collections.Concurrent;
using System.Threading;
using Microsoft.Identity.Client;

namespace Demo.Function
{
    public class HttpTriggerSharepointServices
    {
        private const string APPROVAL_HISTORY_LIST_NAME = "ApprovalHistory";
        private const int BORDER_WIDTH = 1;
        readonly IPnPContextFactory _pnpContextFactory;
        private static ConcurrentDictionary<string, string> _runningTasks = new ConcurrentDictionary<string, string> ();
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
        /// Convert .DWG file to PDF on the fly
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>
        [FunctionName("HttpTrigger1DwgToPdf")]
        public async Task<IActionResult> RunDwgToPdf(
           [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req)
        {
            
            IActionResult fileResult = null;

            #region Parameters
            string outFilePath = string.Empty;
            var destSpFolder = "DwgToPdf";
            var libname = "Documents";
            var flname = "example.dwg";
            var portalKey = "TestPortal";
            var download = false;
            var qry = req.RequestUri.ParseQueryString().GetValues("d");
            var qryd = req.RequestUri.ParseQueryString().GetValues("dwnld");
            var qrylib = req.RequestUri.ParseQueryString().GetValues("lib");
            var qryPortal = req.RequestUri.ParseQueryString().GetValues("pkey");
            var qryDest = req.RequestUri.ParseQueryString().GetValues("dest");

            if (qry != null)
                flname = qry.FirstOrDefault();

            if (qryd != null)
                bool.TryParse(qryd.FirstOrDefault(), out download);

            if (qrylib != null)
                libname = qrylib.FirstOrDefault();

            if (qryPortal != null)
                portalKey = qryPortal.FirstOrDefault();
            if (qryDest != null)
                destSpFolder = qryDest.FirstOrDefault();
            #endregion

            if (_runningTasks.ContainsKey(flname))
                return new JsonResult(new { Status = "In progress" });

            
            Action task = new Action(async () =>
            {

                Logger.LogInformation("C# HTTP trigger function processed a request.");
                var result = new { };

                using (var ctx = await _pnpContextFactory.CreateAsync(portalKey))
                {
                    var destinationLibrary = await ctx.Web.Lists.GetByTitleAsync(destSpFolder, l => l.RootFolder);
                    var shareDocuments = await ctx.Web.Lists.GetByTitleAsync(libname, l => l.RootFolder);

                    var folderContents = await shareDocuments.RootFolder.GetAsync(o => o.Files);
                    var documents = from d in folderContents.Files.Where(o => o.Name.ToLower() == flname.ToLower()).AsEnumerable()
                                    select new
                                    {
                                        d.Name,
                                        d.TimeLastModified,
                                        d.UniqueId
                                    };

                    IFile dwgFile = null;

                    var tmpLocation = System.IO.Path.GetTempPath();

                    foreach (var fl in folderContents.Files.AsEnumerable())
                    {
                        if (Path.GetFileName(fl.Name).Contains(flname))
                        {
                            Logger.LogInformation("Dwg found");
                            Logger.LogInformation(fl.Name);
                            dwgFile = fl;
                            break;
                        }
                    }

                   
                    var tmpFileName = Path.GetFileNameWithoutExtension(flname);
                    var oflName = $"{tmpFileName}.pdf";
                    outFilePath = Path.Combine(tmpLocation, oflName);

                    if (File.Exists(outFilePath))
                    {
                        try
                        {
                            File.Delete(outFilePath);
                        }
                        catch (Exception)
                        {
                            Logger.LogInformation($"Outfile existed");
                            Logger.LogWarning($"couldn't clean outfile - {outFilePath}");
                        }
                    }

                    Logger.LogInformation($"Outfile: {outFilePath}");

                    if (dwgFile != null)
                    {
                        await dwgFile.ListItemAllFields.LoadAsync();

                        var status = dwgFile.ListItemAllFields["ProcessComplete"]?.ToString();
                        
                        if (status == "Y")
                        {
                            Logger.LogInformation($"File[{flname}] was already processed");
                            return;
                        }

                        Logger.LogInformation($"converting file [{dwgFile}] to pdf");
                        using (var ms = new MemoryStream())
                        {
                            var bufferSize = 2 * 1024 * 1024;
                            using (var stream = await dwgFile.GetContentAsync(streamContent: true))
                            {
                                var buffer = new byte[bufferSize];
                                int read;
                                while ((read = await stream.ReadAsync(buffer, 0, buffer.Length)) != 0)
                                {
                                    await ms.WriteAsync(buffer, 0, read);
                                }
                            };

                            ms.Seek(0, SeekOrigin.Begin);

                            Logger.LogInformation($"Downloading file - done");
                            //  Logger.LogInformation($"Downloading file content size - {contentLength}");
                            if (true)
                            {
                                try
                                {
                                    //stream = new FileStream(@"D:\Downloads\sample1.dwg", FileMode.Open, FileAccess.Read);
                                    var opt = new Aspose.CAD.LoadOptions();
                                    var running = true;
                                    // update progress
                                    dwgFile.ListItemAllFields["ProcessingStatus"] = $"{DateTime.Now}|Started";
                                    await dwgFile.ListItemAllFields.UpdateAsync();
                                    
                                    Task statusTask = Task.Factory.StartNew(async () =>
                                    {
                                        while (running)
                                        {
                                            Thread.Sleep(1000 * 60);
                                            dwgFile.ListItemAllFields["ProcessingStatus"] = $"{DateTime.Now}|LastUpdated";
                                            await dwgFile.ListItemAllFields.UpdateAsync();
                                        }
                                    });

                                    using (var img = Aspose.CAD.Image.Load(ms, opt))
                                    {
                                        Aspose.CAD.ImageOptions.CadRasterizationOptions rasterizationOptions = new Aspose.CAD.ImageOptions.CadRasterizationOptions();
                                        rasterizationOptions.PageWidth = img.Width;
                                        rasterizationOptions.PageHeight = img.Height;
                                        rasterizationOptions.AutomaticLayoutsScaling = true;
                                        rasterizationOptions.NoScaling = false;
                                        rasterizationOptions.Margins = new Aspose.CAD.ImageOptions.Margins()
                                        {
                                            Left = 5,
                                            Right = 5,
                                            Bottom = 5,
                                            Top = 5
                                        };

                                        // Create an instance of PdfOptions
                                        Aspose.CAD.ImageOptions.PdfOptions pdfOptions = new Aspose.CAD.ImageOptions.PdfOptions();

                                        //Set the VectorRasterizationOptions property
                                        pdfOptions.VectorRasterizationOptions = rasterizationOptions;

                                        //Export CAD to PDF
                                        using (var os = new MemoryStream())

                                            img.Save(ms, pdfOptions);

                                        ms.Seek(0, SeekOrigin.Begin);
                                        // save to destination sharepoint library
                                        await destinationLibrary.RootFolder.Files.AddAsync(oflName, ms, true);
                                        running = false;
                                        statusTask.Wait();
                                        dwgFile.ListItemAllFields["ProcessingStatus"] = $"{DateTime.Now}|Completed";
                                        dwgFile.ListItemAllFields["ProcessComplete"] = $"Y";
                                        await dwgFile.ListItemAllFields.UpdateAsync();

                                        Logger.LogInformation($"conversion of file [{dwgFile}] to pdf is done");
                                    }
                                }
                                catch (Exception ex)
                                {
                                    Logger.LogError(ex, ex.Message);
                                    throw;
                                }

                                finally
                                {
                                    _runningTasks.TryRemove(flname, out string value);
                                }

                                fileResult = new PhysicalFileResult(outFilePath, "application/pdf")
                                {
                                    FileDownloadName = oflName,
                                };

                               
                            }
                        }
                    }
                    else
                    {
                        fileResult = new NotFoundResult();
                    }
                }
            });

            //if (download)
            //{
            //    Task.Factory.StartNew(task).Wait();
            //    return await Task.FromResult(fileResult);
            //}
           
            Task.Run(task).ContinueWith((t) =>
            {
                Logger.LogInformation("DwgToPdf task is done");
            }).Wait(0);

            await Task.CompletedTask;

            return new JsonResult(new { Status = "Queued"});
        }


        [FunctionName("HttpTrigger1MyFunc")]
        public async Task<IActionResult> Run(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req)
        {
            Logger.LogInformation("C# HTTP trigger function processed a request.");
            var result = new { };

            using (var ctx = await _pnpContextFactory.CreateAsync("Default"))
            {
                var shareDocuments = await ctx.Web.Lists.GetByTitleAsync("Documents", l => l.RootFolder);
                var folderContents = await shareDocuments.RootFolder.GetAsync(o => o.Files);
                var documents = from d in folderContents.Files.AsEnumerable()
                                select new
                                {
                                    d.Name,
                                    d.TimeLastModified,
                                    d.UniqueId
                                };

                IFile pdf = null;
                var tmpLocation = System.IO.Path.GetTempPath();

                foreach (var fl in folderContents.Files.AsEnumerable())
                {
                    if (Path.GetExtension(fl.Name) == ".pdf")
                    {
                        Logger.LogInformation("Pdf found");
                        Logger.LogInformation(fl.Name);
                        pdf = fl;
                        break;
                    }
                }
                var outFilePath = Path.Combine(tmpLocation, "sample-signed-01.pdf");
                Logger.LogInformation($"Outfile: {outFilePath}");
                if (pdf != null)
                {
                    var contents = pdf.GetContentBytes();
                    if (contents.Length > 0)
                    {
                        SignatureHelper.Sign(contents,
                            outFilePath,
                            @"D:\Azure Certifications\az204\Tutorials\myfunc\Certificates\PnP.Core.SDK.AzureFunctionSample.cer",
                            "Approved",
                            "NOIDA, INDIA"
                            );

                        return new PhysicalFileResult(outFilePath, "application/pdf");
                    }
                }
                else
                {
                    return new JsonResult(documents);
                }

            }

            return new NotFoundResult();
        }

        /// <summary>
        /// Adds approval history of a office 365 docx from a sharepoint library, run against TestPortal site
        /// </summary>
        /// <param name="req"></param>
        /// <returns></returns>

        [FunctionName("HttpTrigger1MyFunc2")]
        public async Task<IActionResult> Run2(
            [HttpTrigger(AuthorizationLevel.Function, "get", "post", Route = null)] HttpRequestMessage req)
        {
            Logger.LogInformation("C# HTTP trigger function processed a request.");

            var result = new { };

            var docid = string.Empty;
            var libname = "Documents";
            var flname = "Document.docx";
            var download = false;
            var qry = req.RequestUri.ParseQueryString().GetValues("d");
            var qryd = req.RequestUri.ParseQueryString().GetValues("dwnld");
            var qrylib = req.RequestUri.ParseQueryString().GetValues("lib");
            var qDocId = req.RequestUri.ParseQueryString().GetValues("docid");
            if (qry != null)
                flname = qry.FirstOrDefault();

            if (qryd != null)
                bool.TryParse(qryd.FirstOrDefault(), out download);

            if (qrylib != null)
                libname = qrylib.FirstOrDefault();

            if (qDocId != null)
                docid = qDocId.FirstOrDefault();

            using (var ctx = await _pnpContextFactory.CreateAsync("TestPortal"))
            {
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
                        Logger.LogInformation($"docid:{docid}");
                        Logger.LogInformation("Docx found");
                        Logger.LogInformation(fl.Name);
                        docx = fl;
                        // return new NotFoundResult();
                    }
                }



                if (docx == null || string.IsNullOrEmpty(docid))
                    return new NotFoundResult();

                else
                {

                    string viewXml = @"<View>
                    <Query>
                      <Where>
                        <Eq>
                          <FieldRef Name='DMSID'/>
                          <Value Type='text'>" + docid + @"</Value>
                        </Eq>
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

                        table.SetAttribute(attrib);

                        TableProperties tblProp = CreateTableProperties();

                        // Append the TableProperties object to the empty table.
                        table.AppendChild<TableProperties>(tblProp);

                        TableRow trHead = new TableRow();

                        CreateCell("Role", trHead, true);
                        CreateCell("Name", trHead, true);
                        CreateCell("Date Of Approval", trHead, true);
                        CreateCell("Comment", trHead, true);
                        CreateCell("Action", trHead, true);
                        table.Append(trHead);

                        foreach (var item in historyItems)
                        {
                            var val1 = Convert.ToString(item["Role"]);
                            var val2 = Convert.ToString(item["UserName"]);
                            var val3 = Convert.ToDateTime(item["Created"]).ToString("dd-MMM-yyyy");
                            var val4 = Convert.ToString(item["Comment"]);
                            var val5 = Convert.ToString(item["Action"]);

                            TableRow tr = new TableRow();

                            TableRowProperties trProp = new TableRowProperties(new TableRowHeight
                            {
                                HeightType = new EnumValue<HeightRuleValues>(HeightRuleValues.Auto),
                            });

                            CreateCell(val1, tr);
                            CreateCell(val2, tr);
                            CreateCell(val3, tr);
                            CreateCell(val4, tr, false, 2160);
                            CreateCell(val5, tr);

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

                        table.SetAttribute(attrib);

                        TableProperties tblProp = CreateTableProperties();

                        // Append the TableProperties object to the empty table.
                        table.AppendChild<TableProperties>(tblProp);

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
