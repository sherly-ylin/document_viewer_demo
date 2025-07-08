using document_viewer_demo.Models;
using Microsoft.AspNetCore.Mvc;
using System.Data;
using System.Diagnostics;
using Newtonsoft.Json;
using TXTextControl;
using TXTextControl.DocumentServer;
using Microsoft.Data.SqlClient;
using System.Threading.Tasks;
using System.Text.RegularExpressions;

namespace document_viewer_demo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        List<int> orderIds = new List<int> { 7259, 7261, 7262, 7264, 7266 };
        // List<int> orderIds = new List<int> { 7262, 7264 };
        // int bundleOrderId = 7262;
        bool testBundle = true;
        string connectionString = "Server=192.168.20.97;Database=SalesChain0602_MS_MN;User Id=ylin;Password=9244@Wahg;TrustServerCertificate=True;";


        // private List<int> pageLengths { get; set; } = new List<int>();
        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public async Task<IActionResult> Index()
        {
            try
            {
                // string sessionKey = $"merged_document_{string.Join("_", orderIds)}";
                string sessionKey = $"sc_document_{DateTime.Now}";
                Console.WriteLine(sessionKey);

                Console.WriteLine("=====> Session Key: " + sessionKey);
                // Try to get from session first
                string docBase64 = HttpContext.Session.GetString(sessionKey);

                if (string.IsNullOrEmpty(docBase64))
                {
                    _logger.LogInformation("Document not found in session, generating new document");
                    // Generate and store in session

                    // docBase64 = Convert.ToBase64String(System.IO.File.ReadAllBytes("Documents/OBDWC - Copy.docx"));
                    docBase64 = GetDocumentBytes("Documents/OBDWC - Copy.docx", StreamType.WordprocessingML);
                    // docBase64 = await LoadTemplateAndMergeMultipleOrders(orderIds);
                    // docBase64 = LoadTemplateAndMergeMultipleOrders(orderIds);

                    HttpContext.Session.SetString(sessionKey, docBase64);
                }
                else
                {
                    _logger.LogInformation("Document retrieved from session cache");
                }

                ViewBag.HasDocument = true;
                ViewBag.DocumentData = docBase64;
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
        public IActionResult ClearCache()
        {
            string sessionKey = $"merged_document_{string.Join("_", orderIds)}";

            HttpContext.Session.Remove(sessionKey);
            _logger.LogInformation("Document cache cleared from session");

            return RedirectToAction("Index");
        }

        void ConvertToMergeFields(ServerTextControl tx)
        {
            string pattern = @"\{\{(.*?)\}\}";
            var regex = new Regex(pattern);

            // Move to the start of the document
            tx.Select(0, 0);
            var count = 0;
            Console.WriteLine(tx.Text);
            Console.WriteLine($"Input Posisition {tx.InputPosition}");
            while (regex.IsMatch(tx.Text) && count < 2)
            {
                count++;
                Console.WriteLine("Match {}");
                var match = regex.Match(tx.Text);
                Console.WriteLine(match);
                if (match.Success)
                {
                    Console.WriteLine("match success");
                    string placeholder = match.Groups[0].Value; // e.g., {{OR.OrderID}}
                    string fieldName = match.Groups[1].Value;    // e.g., OR.OrderID

                    Console.WriteLine($"{placeholder}-{fieldName}");

                    int start = tx.Text.IndexOf(placeholder);
                    Console.WriteLine($"start {start}/{placeholder.Length}");
                    if (start < 0) break;

                    tx.Select(start, placeholder.Length);
                    tx.Clear(); // remove old text
                                // tx.Select(0, 0);
                    Console.WriteLine($"Input Posisition {tx.InputPosition}");

                    tx.Select(start, start);
                    Console.WriteLine($"- Select Start / Input Posisition {tx.InputPosition}");

                    ApplicationField field
                    = new ApplicationField(ApplicationFieldFormat.MSWord,
                    "MERGEFIELD",
                    fieldName,
                    [fieldName]);
                    tx.ApplicationFields.Add(field);

                    // TXTextControl.DocumentServer.Fields.MergeField mergeField =
                    // new TXTextControl.DocumentServer.Fields.MergeField()
                    // {
                    //     Text = fieldName,
                    //     Name = fieldName,
                    //     TextBefore = ""
                    // };
                    // tx.ApplicationFields.Add(mergeField.ApplicationField);
                }
                // break;
            }
        }

        public IActionResult DownloadSelectedPages(int[] pageNumbers, string fileName, string sessionKey)
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
                string docBase64 = HttpContext.Session.GetString(sessionKey);
                if (string.IsNullOrEmpty(docBase64))
                {
                    return BadRequest("Document no longer available in session. Please refresh the page to regenerate the document.");
                }

                _logger.LogInformation($"Downloading selected pages: {string.Join(", ", pageNumbers)}");

                byte[] documentBytes = Convert.FromBase64String(docBase64);
                byte[] selectedPagesBytes = ExtractSelectedPages(documentBytes, pageNumbers);
                byte[] pdfBytes = ConvertToPdf(selectedPagesBytes);

                // string fileName = $"selected_pages_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
                return File(pdfBytes, "application/pdf", $"{fileName}.pdf");
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error downloading selected pages");
                return StatusCode(500, "Error processing selected pages: " + ex.Message);
            }
        }

        // public IActionResult DownloadDocumentAsPdf(string sessionKey)
        // {
        //     try
        //     {
        //         if (string.IsNullOrEmpty(sessionKey))
        //         {
        //             return BadRequest("Session key is missing");
        //         }

        //         // Get the document from session
        //         string docBase64 = HttpContext.Session.GetString(sessionKey);
        //         if (string.IsNullOrEmpty(docBase64))
        //         {
        //             return BadRequest("Document no longer available in session. Please refresh the page to regenerate the document.");
        //         }

        //         _logger.LogInformation("Downloading full document as PDF");

        //         byte[] documentBytes = Convert.FromBase64String(docBase64);
        //         byte[] pdfBytes = ConvertToPdf(documentBytes);

        //         string fileName = $"merged_orders_{DateTime.Now:yyyyMMdd_HHmmss}.pdf";
        //         return File(pdfBytes, "application/pdf", fileName);
        //     }
        //     catch (Exception ex)
        //     {
        //         _logger.LogError(ex, "Error downloading PDF");
        //         return StatusCode(500, "Error generating PDF: " + ex.Message);
        //     }
        // }

        public void GetPageNumbers(ServerTextControl tx)
        {
            PageCollection pages = tx.GetPages();
            for (int i = 0; i < pages.Count; i++)
            {
                Console.WriteLine($"page {i}/ {pages[i].Number}");
            }
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
                    PageCollection pages = sourceTx.GetPages();
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
                        // Console.WriteLine($"Extracting page {pageNumbers[i]}: Start={pageStartPositions[pageNumbers[i] - 1]}, End={pageStartPositions[pageNumbers[i]] - pageStartPositions[pageNumbers[i] - 1] - 1}, Length={page.Length}");

                        // sourceTx.Select(pageStartPositions[pageNumbers[i] - 1], pageStartPositions[pageNumbers[i]] - pageStartPositions[pageNumbers[i] - 1]);
                        // sourceTx.Select(page, pageStartPositions[pageNumbers[i]] - pageStartPositions[pageNumbers[i] - 1]);

                        byte[] pageContent;
                        sourceTx.Selection.Save(out pageContent, BinaryStreamType.InternalUnicodeFormat);
                        // Console.WriteLine($"Extracted page {pageNumbers[i]} content length: {pageStartPositions[pageNumbers[i]] - pageStartPositions[pageNumbers[i] - 1]}");

                        targetTx.Append(pageContent, BinaryStreamType.InternalUnicodeFormat, AppendSettings.None);
                    }
                    var index = targetTx.Find("\f", -1, FindOptions.Reverse);
                    // Console.WriteLine($"Selecttion complete == Page break found at index: {index}");

                    if (index > 0)
                    {
                        // Clear the last page break if it exists
                        targetTx.Select(index, 1);
                        targetTx.Clear();
                        // Console.WriteLine("Cleared last char");
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

        private string GetDocumentBytes(string filePath, StreamType streamType)
        {
            using (ServerTextControl tx = new ServerTextControl())
            {
                tx.Create();

                LoadTemplate(tx, filePath, streamType);
                // PageCollection pages = tx.GetPages();
                // for (int i = 0; i < pages.Count; i++)
                // {
                //     Console.WriteLine($"page {i}/ {pages.GetItem(i).Start}/ {pages.GetItem(i).Length}/ {pages.GetItem(i).Footer}");
                // }
                tx.Select(200, tx.Text.Length - 200);
                tx.Clear();

                using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                {
                    string jsonData = System.IO.File.ReadAllText("Documents/jsonData.json");
                    Console.WriteLine(jsonData);
                    mailMerge.MergeJsonData(jsonData);

                }
                ConvertToMergeFields(tx);
                // SectionCollection sections = tx.Sections;
                // Console.WriteLine("number of sections: " + sections.Count);

                // var loadSettings = new LoadSettings
                // {
                //     ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                //     LoadSubTextParts = true
                // };

                // tx.Load(filePath, streamType, loadSettings);

                byte[] bytes;
                tx.Save(out bytes, BinaryStreamType.InternalUnicodeFormat);
                return Convert.ToBase64String(bytes);
            }
        }
        private void LoadTemplate(ServerTextControl tx, string fp, StreamType streamType = StreamType.AllFormats)
        {

            fp = fp.Trim();
            string formatStr = fp.Substring(fp.LastIndexOf('.'));
            streamType = formatStr switch
            {
                ".pdf" => StreamType.AdobePDF,
                ".rtf" => StreamType.RichTextFormat,
                ".tx" => StreamType.InternalUnicodeFormat,
                _ => StreamType.WordprocessingML
            };
            Console.WriteLine(streamType);

            var loadSettings = new LoadSettings
            {
                ApplicationFieldFormat = ApplicationFieldFormat.MSWord
                // LoadSubTextParts = true
            };

            tx.Load(fp, streamType, loadSettings);
        }
        private async Task<string> LoadTemplateAndMergeMultipleOrders(List<int> orderIds)
        {
            Console.WriteLine("=== Merging multiple orders: " + string.Join(", ", orderIds));
            using (ServerTextControl masterTx = new ServerTextControl())
            {
                masterTx.Create();
                var breakInd = -1;

                for (int i = 0; i < orderIds.Count; i++)
                {
                    Console.WriteLine("=====Processing OrderId: " + orderIds[i]);

                    using (ServerTextControl tx = new ServerTextControl())
                    {
                        SNOrder dbOrder = GetMappedOrderObj(orderIds[i]);
                        string jsonData = await GetOrderDataJsonFromDb(orderIds[i], true);

                        // Console.WriteLine(JsonConvert.SerializeObject(dbOrder));
                        tx.Create();

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
                        using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                        {
                            mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                            // mailMerge.MergeObject(dbOrder);
                            mailMerge.MergeJsonData(jsonData);
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
                breakInd = masterTx.Find("\f", -1, FindOptions.Reverse);
                masterTx.Select(breakInd, 1);
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

        public string GetQuery()
        {
            var query = "";

            return query;
        }

        public async Task<byte[]> GetDocumentFromDB(int fileId)
        {
            using var connection = new SqlConnection(connectionString);
            using var command = new SqlCommand("SELECT * FROM Documents WHERE FileId = @Id", connection);
            command.Parameters.AddWithValue("@Id", fileId);
            await connection.OpenAsync();
            var result = await command.ExecuteScalarAsync();
            return result as byte[];
        }

        public async Task<string> GetOrderDataJsonFromDb(int orderId, bool byBundle = false)
        {
            Console.WriteLine("Retrieving order info from database for OrderId: " + orderId);

            string jsonRes = string.Empty;
            using (var conn = new SqlConnection(connectionString))
            {
                await conn.OpenAsync();
                try
                {
                    // ,BundleQuantity, PerBundleQuantity, BundleOrder, ol.SeqNbr
                    var query = @"SELECT o.OrderID, o.CustomerName, o.SubTotalAmount TotalSellPrice,  o.DTCreated, 
		                            CONCAT(
                                        TRIM(o.BillingAddress1), ' ', TRIM(o.BillingAddress2), ' ' , 
                                        TRIM(o.BillingCity), ', ', TRIM(o.BillingState),' ' , TRIM(o.BillingPostalCode)
                                    ) AS BillingAddress,
                                    (SELECT ol.OrderLineId, ol.Model, ol.SellPrice, ol.Quantity, ol.LineTotal, BundleID
                                    FROM SNOrderLine ol WHERE o.OrderId = ol.OrderId
                                    ORDER BY ol.OrderID, ol.BundleID
                                    FOR JSON PATH) AS OrderLines
                                FROM SNOrder o 
                                WHERE o.OrderId = @OrderId
                                FOR JSON PATH;";
                    if (byBundle)
                    {
                        query = @"SELECT o.OrderID, o.CustomerName, o.SubTotalAmount AS TotalSellPrice, o.DTCreated,
                                        CONCAT(
                                            TRIM(o.BillingAddress1), ' ', TRIM(o.BillingAddress2), ' ', 
                                            TRIM(o.BillingCity), ', ', TRIM(o.BillingState), ' ', 
                                            TRIM(o.BillingPostalCode)
                                        ) AS BillingAddress,
                                        (SELECT ol.BundleID, SUM(ol.LineTotal) AS BundleTotal,
                                            (SELECT ol2.OrderLineId, TRIM(ol2.Model) AS Model, ol2.SellPrice, ol2.Quantity, ol2.LineTotal
                                                FROM SNOrderLine ol2
                                                WHERE ol2.OrderId = o.OrderId AND ol2.BundleID = ol.BundleID
                                                FOR JSON PATH
                                            ) AS OrderLines
                                            FROM SNOrderLine ol
                                            WHERE ol.OrderId = o.OrderId
                                            GROUP BY ol.BundleID
                                            FOR JSON PATH
                                        ) AS OrderBundle
                                        FROM SNOrder o
                                        WHERE o.OrderId = 7262
                                        FOR JSON PATH
                                        ";
                    }
                    var command = new SqlCommand(query, conn);
                    command.Parameters.AddWithValue("@OrderId", orderId);

                    using var reader = await command.ExecuteReaderAsync();

                    if (await reader.ReadAsync())
                    {
                        jsonRes = reader.GetString(0);
                        Console.WriteLine("jsonRes:");
                        Console.WriteLine(jsonRes);
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

            return jsonRes;
        }

        public DataTable GetOrderDataFromDb(int orderId)
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
                    var query = @"SELECT o.OrderID, o.CustomerName, o.BillingAddress1, o.BillingAddress2, o.BillingCity, o.BillingState, o.BillingPostalCode, o.DTCreated,
                                ol.OrderLineId, ol.Model, ol.BundleID, ol.SellPrice, ol.Quantity, ol.LineTotal FROM SNOrder o 
                                JOIN SNOrderLine ol on o.OrderId = ol.OrderId
                                WHERE o.OrderId = @OrderId
                                ORDER BY ol.BundleID, Model";

                    var command = new SqlCommand(query, conn);

                    command.Parameters.AddWithValue("@OrderId", orderId);
                    using (var reader = command.ExecuteReader())
                    {
                        resultTable.Load(reader);
                        Console.WriteLine("Total results rows: " + resultTable.Rows.Count);
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

            return resultTable;
        }
        public SNOrder GetMappedOrderObj(int orderId)
        {
            Console.WriteLine("Retrieving order info from database for OrderId: " + orderId);
            var order = new SNOrder();

            DataTable resultTable = GetOrderDataFromDb(orderId);

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

            return order;
        }
        public string GetMappedOrderJson(int orderId)
        {
            DataTable resultTable = GetOrderDataFromDb(orderId);
            string jsonData = JsonConvert.SerializeObject(resultTable);
            return jsonData;
        }

        List<OrderBundle> SplitOrderIntoBundles(SNOrder order)
        {
            return order.OrderLines
                .GroupBy(item => item.BundleID)
                .Select(group => new OrderBundle
                {
                    BundleID = group.Key,
                    CustomerName = order.CustomerName,
                    BillingAddress = order.BillingAddress,
                    DTCreated = order.DTCreated,
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