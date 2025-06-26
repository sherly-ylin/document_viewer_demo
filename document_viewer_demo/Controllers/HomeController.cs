using document_viewer_demo.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;
using Newtonsoft.Json;
using TXTextControl;
using TXTextControl.DocumentServer;
using Microsoft.Data.SqlClient;

namespace document_viewer_demo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        // List<int> orderIds = new List<int> { 7259, 7261, 7262, 7264, 7266 };
        List<int> orderIds = new List<int> { 7262, 7264 };
        // int bundleOrderId = 7262;
        bool testBundle = true;

        // private List<int> pageLengths { get; set; } = new List<int>();
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            try
            {
                string sessionKey = $"merged_document_{string.Join("_", orderIds)}";


                Console.WriteLine("=====> Session Key: " + sessionKey);
                // Try to get from session first
                string generatedDocumentBase64 = HttpContext.Session.GetString(sessionKey);

                if (string.IsNullOrEmpty(generatedDocumentBase64))
                {
                    _logger.LogInformation("Document not found in session, generating new document");
                    // Generate and store in session
                    // generatedDocumentBase64 = LoadDocument("Documents/sample.docx", StreamType.WordprocessingML);

                    generatedDocumentBase64 = LoadTemplateAndMergeMultipleOrders(orderIds);
                    generatedDocumentBase64 = LoadTemplateAndMergeMultipleOrders(orderIds);

                    HttpContext.Session.SetString(sessionKey, generatedDocumentBase64);
                }
                else
                {
                    _logger.LogInformation("Document retrieved from session cache");
                }

                ViewBag.HasDocument = true;
                ViewBag.DocumentData = generatedDocumentBase64;
                ViewBag.SessionKey = sessionKey;
                ViewBag.DocumentName = $"Merged_Orders_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            }
            catch (Exception ex)
            {
                ViewBag.HasDocument = false;
                ViewBag.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Error processing document template");
            }

            return View();
        }

        private string LoadDocument(string filePath, StreamType streamType)
        {
            using (ServerTextControl tx = new ServerTextControl())
            {
                tx.Create();

                // Load the template
                var loadSettings = new LoadSettings
                {
                    ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                    LoadSubTextParts = true
                };
                tx.Load(filePath, streamType, loadSettings);
                SectionCollection sections = tx.Sections;
                Console.WriteLine("number of sections: " + sections.Count);

                byte[] bytes;
                tx.Save(out bytes, BinaryStreamType.InternalUnicodeFormat);
                return Convert.ToBase64String(bytes);
            }
        }
        public IActionResult DownloadSelectedPages(int[] pageNumbers, string sessionKey)
        {
            try
            {
                if (pageNumbers == null || pageNumbers.Length == 0)
                {
                    return BadRequest("No pages selected");
                }

                if (string.IsNullOrEmpty(sessionKey))
                {
                    return BadRequest("Session key is missing");
                }

                // Get the document from session
                string generatedDocumentBase64 = HttpContext.Session.GetString(sessionKey);
                if (string.IsNullOrEmpty(generatedDocumentBase64))
                {
                    return BadRequest("Document no longer available in session. Please refresh the page to regenerate the document.");
                }

                _logger.LogInformation($"Downloading selected pages: {string.Join(", ", pageNumbers)}");

                byte[] documentBytes = Convert.FromBase64String(generatedDocumentBase64);
                byte[] selectedPagesBytes = ExtractSelectedPages(documentBytes, pageNumbers);
                byte[] pdfBytes = ConvertToPdf(selectedPagesBytes);

                string fileName = $"selected_pages_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                return File(pdfBytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading selected pages");
                return StatusCode(500, "Error processing selected pages: " + ex.Message);
            }
        }

        public IActionResult DownloadDocumentAsPdf(string sessionKey)
        {
            try
            {
                if (string.IsNullOrEmpty(sessionKey))
                {
                    return BadRequest("Session key is missing");
                }

                // Get the document from session
                string generatedDocumentBase64 = HttpContext.Session.GetString(sessionKey);
                if (string.IsNullOrEmpty(generatedDocumentBase64))
                {
                    return BadRequest("Document no longer available in session. Please refresh the page to regenerate the document.");
                }

                _logger.LogInformation("Downloading full document as PDF");

                byte[] documentBytes = Convert.FromBase64String(generatedDocumentBase64);
                byte[] pdfBytes = ConvertToPdf(documentBytes);

                string fileName = $"merged_orders_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                return File(pdfBytes, "application/pdf", fileName);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading PDF");
                return StatusCode(500, "Error generating PDF: " + ex.Message);
            }
        }

        // Helper method to clear session cache (useful for testing or manual refresh)
        public IActionResult ClearCache()
        {
            string sessionKey = $"merged_document_{string.Join("_", orderIds)}";

            HttpContext.Session.Remove(sessionKey);
            _logger.LogInformation("Document cache cleared from session");

            return RedirectToAction("Index");
        }
        private byte[] ExtractSelectedPages(byte[] documentBytes, int[] pageNumbers)
        {
            using (ServerTextControl sourceTx = new ServerTextControl())
            {
                sourceTx.Create();
                sourceTx.Load(documentBytes, BinaryStreamType.InternalUnicodeFormat);

                using (ServerTextControl targetTx = new ServerTextControl())
                {
                    targetTx.Create();

                    // Sort page numbers to maintain order
                    Array.Sort(pageNumbers);
                    var pages = sourceTx.GetPages();
                    var pageLengths = Enumerable.Range(0, pages.Count)
                        .Select(i => pages.GetItem(i).Length).ToList();

                    // Calculate page start positions: sum of lengths of previous pages + number of previous pages
                    var pageStartPositions = new List<int> { 0 };
                    int currPos = pageLengths[0];
                    sourceTx.Append("\f", StringStreamType.PlainText, AppendSettings.None); // For calculation purposes

                    for (int i = 1; i <= pageLengths.Count; i++)
                    {
                        var indexPageBreak = sourceTx.Find("\f", pageStartPositions[i - 1], FindOptions.MatchWholeWord);
                        Console.WriteLine($"Page break found at index: {indexPageBreak}");
                        pageStartPositions.Add(indexPageBreak + 1);
                    }


                    Console.WriteLine("Total pages in document: " + pages.Count);
                    Console.WriteLine("Page lengths: " + string.Join(", ", pageLengths));
                    Console.WriteLine("Page start positions: " + string.Join(", ", pageStartPositions));

                    for (int i = 0; i < pageNumbers.Length; i++)
                    {
                        if (pageNumbers[i] < 1 || pageNumbers[i] > pages.Count)
                        {
                            continue;
                        }

                        var page = pages.GetItem(pageNumbers[i] - 1); // Pages are 0-indexed
                        Console.WriteLine($"Extracting page {pageNumbers[i]}: Start={pageStartPositions[pageNumbers[i] - 1]}, End={pageStartPositions[pageNumbers[i]] - pageStartPositions[pageNumbers[i] - 1] - 1}, Length={page.Length}");

                        sourceTx.Select(pageStartPositions[pageNumbers[i] - 1], pageStartPositions[pageNumbers[i]] - pageStartPositions[pageNumbers[i] - 1]);

                        byte[] pageContent;
                        sourceTx.Selection.Save(out pageContent, BinaryStreamType.InternalUnicodeFormat);
                        Console.WriteLine($"Extracted page {pageNumbers[i]} content length: {pageStartPositions[pageNumbers[i]] - pageStartPositions[pageNumbers[i] - 1]}");

                        targetTx.Append(pageContent, BinaryStreamType.InternalUnicodeFormat, AppendSettings.None);
                    }
                    var index = targetTx.Find("\f", -1, FindOptions.Reverse);
                    Console.WriteLine($"Selecttion complete == Page break found at index: {index}");
                    if (index > 0)
                    {
                        // Clear the last page break if it exists
                        targetTx.Select(index, 1);
                        targetTx.Clear();
                        Console.WriteLine("Cleared last char");
                    }

                    // Save the extracted pages
                    byte[] result;
                    targetTx.Save(out result, BinaryStreamType.InternalUnicodeFormat);
                    return result;
                }
            }
        }
        private byte[] ConvertToPdf(byte[] documentBytes)
        {
            using (ServerTextControl tx = new ServerTextControl())
            {
                tx.Create();
                tx.Load(documentBytes, BinaryStreamType.InternalUnicodeFormat);

                byte[] pdfBytes;
                tx.Save(out pdfBytes, BinaryStreamType.AdobePDF);
                return pdfBytes;
            }
        }

        private string LoadTemplateAndMergeMultipleOrders(List<int> orderIds)
        {
            Console.WriteLine("=== Merging multiple orders: " + string.Join(", ", orderIds));
            using (ServerTextControl masterTx = new ServerTextControl())
            {
                masterTx.Create();

                for (int i = 0; i < orderIds.Count; i++)
                {
                    Console.WriteLine("Processing OrderId: " + orderIds[i]);
                    using (ServerTextControl tx = new ServerTextControl())
                    {
                        tx.Create();

                        // Load the template
                        var loadSettings = new LoadSettings
                        {
                            ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                            LoadSubTextParts = true
                        };

                        if (testBundle)
                        {
                            tx.Load("Documents/template_order_bundle.docx", StreamType.WordprocessingML, loadSettings);

                        }
                        else
                        {
                            tx.Load("Documents/template_order.docx", StreamType.WordprocessingML, loadSettings);
                        }

                            SNOrder dbOrder = GetOrderFromDb(orderIds[i]);
                            Console.WriteLine(JsonConvert.SerializeObject(dbOrder));
                        if (testBundle) { }
                        
                            using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                        {
                            mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                            mailMerge.MergeObject(dbOrder);
                        }
                        
                        byte[] bytes;
                        tx.Save(out bytes, BinaryStreamType.InternalUnicodeFormat);

                        // SectionCollection sections = masterTx.Sections;
                        // Console.WriteLine("number of sections masterTX: " + sections.Count);

                        Console.WriteLine("Appending document for OrderId: " + orderIds[i]);
                        masterTx.Append(bytes, BinaryStreamType.InternalUnicodeFormat, AppendSettings.None);
                        Console.WriteLine("Appending page break");
                        masterTx.Append("\f", StringStreamType.PlainText, AppendSettings.None);


                        // sections = masterTx.Sections;
                        // Console.WriteLine("number of sections after appending: " + sections.Count);
                        Console.WriteLine("number of pages: " + masterTx.Pages);
                    }
                }

                //Remove the last page break
                var index = masterTx.Find("\f", -1, FindOptions.Reverse);
                masterTx.Select(index, 1);
                masterTx.Clear();

                // Save the merged document to a byte array
                byte[] documentBytes;
                var saveSettings = new SaveSettings
                {
                    CreatorApplication = "Document Viewer Demo"
                };

                masterTx.Save(out documentBytes, BinaryStreamType.InternalUnicodeFormat, saveSettings);

                return Convert.ToBase64String(documentBytes);
            }
        }

        // Keep your existing GetOrderFromDb method unchanged
        public SNOrder GetOrderFromDb(int orderId, bool byBundle = false)
        {
            Console.WriteLine("Retrieving order info from database for OrderId: " + orderId);
            var order = new SNOrder();

            string connectionString = "Server=192.168.20.97;Database=SalesChain0602_MS_MN;User Id=ylin;Password=9244@Wahg;TrustServerCertificate=True;";
            DataTable resultTable = new DataTable();

            using (var conn = new Microsoft.Data.SqlClient.SqlConnection(connectionString))
            {
                conn.Open();
                try
                {

                    if (byBundle)
                    {
                        var query = @"SELECT * FROM SNOrder o 
                                JOIN SNOrderLine ol on o.OrderId = ol.OrderId
                                WHERE o.OrderId = @OrderId
                                ORDER BY ol.BundleID, Model";
                        var cmd = new SqlCommand(query, conn);

                        cmd.Parameters.AddWithValue("@OrderId", orderId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            resultTable.Load(reader);
                        }

                        // For each distinct BundleID in resultTable
                        var distinctBundleIds = resultTable.AsEnumerable()
                            .Select(r => r.Field<int>("BundleID"))
                            .Distinct();

                        foreach (var bundleId in distinctBundleIds)
                        {
                            Console.WriteLine("Processing BundleID: " + bundleId);
                            order.OrderBundles.Add(new OrderBundle
                            {
                                BundleID = bundleId
                            });
                        }
                    }
                    else
                    {
                        var query = @"SELECT * FROM SNOrder o 
                                JOIN SNOrderLine ol on o.OrderId = ol.OrderId
                                WHERE o.OrderId = @OrderId
                                ORDER BY ol.BundleID, Model";
                        var cmd = new SqlCommand(query, conn);

                        cmd.Parameters.AddWithValue("@OrderId", orderId);
                        using (var reader = cmd.ExecuteReader())
                        {
                            resultTable.Load(reader);
                            Console.WriteLine("Total results rows: " + resultTable.Rows.Count);
                        }
                        if (resultTable.Rows.Count > 0)
                        {
                            var row = resultTable.Rows[0];
                            order.OrderID = Convert.ToInt32(row["OrderID"]);
                            order.CustomerName = row["CustomerName"].ToString();
                            order.BillingAddress = row["BillingAddress1"].ToString() + ", " +
                                (string.IsNullOrEmpty(row["BillingAddress2"].ToString()) ? "" : row["BillingAddress2"].ToString() + ", ") +
                                row["BillingCity"].ToString() + ", " +
                                row["BillingState"].ToString() + " " +
                                row["BillingPostalCode"].ToString();
                            order.DTCreated = Convert.ToDateTime(row["DTCreated"]);

                            foreach (DataRow itemRow in resultTable.Rows)
                            {
                                order.OrderLines.Add(new OrderLine
                                {
                                    OrderLineID = Convert.ToInt32(itemRow["OrderLineID"]),
                                    BundleID = Convert.ToInt32(itemRow["BundleID"]),
                                    Model = itemRow["Model"].ToString().TrimEnd(),
                                    Quantity = Convert.ToInt32(itemRow["Quantity"]),
                                    SellPrice = Convert.ToDecimal(itemRow["SellPrice"]),
                                    LineTotal = Convert.ToDecimal(itemRow["LineTotal"])
                                });
                            }
                        }
                    }

                }
                catch (Exception ex)
                {
                    throw new Exception("Error retrieving order info: " + ex.Message, ex);
                }
                finally
                {
                    conn.Close();
                }
            }

            return order;
        }

        List<OrderBundle> SplitOrderIntoBundles(SNOrder order)
        {
            return order.OrderLines
                .GroupBy(item => item.BundleID)
                .Select(group => new OrderBundle
                {
                    BundleID = group.Key,
                    CustomerName = order.CustomerName,
                    ShippingAddress = order.BillingAddress,
                    OrderDate = order.DTCreated,
                    OrderLines = group.ToList()
                })
                .ToList();
        }
        public IActionResult Privacy()
        {
            return View();
        }
    }
}