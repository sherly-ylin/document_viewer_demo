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
using Newtonsoft.Json.Linq;
using Microsoft.IdentityModel.Tokens;
using System.Security.Cryptography.X509Certificates;
namespace document_viewer_demo.Controllers
{
    public class ESignController : Controller
    {
        private readonly ILogger<ESignController> _logger;
        private const string CertificatePath = "Certificates/my_certificate.pfx";
        private const string CertificatePassword = "123123123"; // ⚠️ Consider securing this outside of source code

        string connectionString = "Server=192.168.20.97;Database=SalesChain0602_MS_MN;User Id=ylin;Password=9244@Wahg;TrustServerCertificate=True;";


        // private List<int> pageLengths { get; set; } = new List<int>();
        public ESignController(ILogger<ESignController> logger)
        {
            _logger = logger;
        }

        public async Task<IActionResult> Index()
        {
            try
            {
                _logger.LogInformation("Document not found in session, generating new document");
                // Generate and store in session
                string docBase64 = "";

                // docBase64 = Convert.ToBase64String(System.IO.File.ReadAllBytes("Documents/OBDWC - Copy.docx"));
                docBase64 = await GetDocumentBytes("Documents/signature.tx", StreamType.InternalFormat);
                // docBase64 = await LoadTemplateAndMergeMultipleOrders(orderIds);
                // docBase64 = LoadTemplateAndMergeMultipleOrders(orderIds);

                ViewBag.HasDocument = true;
                ViewBag.DocumentData = docBase64;
                ViewBag.DocumentName = $"Signature_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            }
            catch (Exception ex)
            {
                ViewBag.HasDocument = false;
                ViewBag.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Error processing document template");
            }

            return View();
        }

        [HttpPost]
        public string SignDocument()
        {
            // read the payload
            Stream inputStream = Request.InputStream;
            inputStream.Position = 0;

            StreamReader str = new StreamReader(inputStream);
            string sBuf = str.ReadToEndAsync().Result;

            // retrieve the signature data from the payload
            TXTextControl.Web.MVC.DocumentViewer.Models.SignatureData data =
              JsonConvert.DeserializeObject
              <TXTextControl.Web.MVC.DocumentViewer.Models.SignatureData>(sBuf);

            byte[] bPDF;

            // create temporary ServerTextControl
            using (TXTextControl.ServerTextControl tx = new TXTextControl.ServerTextControl())
            {
                tx.Create();

                // load the document
                tx.Load(Convert.FromBase64String(data.SignedDocument.Document),
                  TXTextControl.BinaryStreamType.InternalUnicodeFormat);

                // create a certificate
                X509Certificate2 cert = new X509Certificate2(
                       Server.MapPath("~/App_Data/textcontrolself.pfx"), "123");

                // assign the certificate to the signature fields
                TXTextControl.SaveSettings saveSettings = new TXTextControl.SaveSettings()
                {
                    CreatorApplication = "TX Text Control Sample Application",
                    SignatureFields = new DigitalSignature[] {
        new TXTextControl.DigitalSignature(cert, null, "txsign"),
        new TXTextControl.DigitalSignature(cert, null, "txsign_init")}
                };

                // save the document as PDF
                tx.Save(out bPDF, TXTextControl.BinaryStreamType.AdobePDFA, saveSettings);
            }

            // return as Base64 encoded string
            return Convert.ToBase64String(bPDF);
        }

        // [HttpPost("SignDocument")]
        // public IActionResult SignDocument([FromBody] SignatureData signatureData)
        // {
        //     if (signatureData?.SignedDocument?.Document == null || string.IsNullOrWhiteSpace(signatureData.UniqueId))
        //     {
        //         return BadRequest("Invalid signature data.");
        //     }

        //     try
        //     {
        //         byte[] signedDocumentBytes = Convert.FromBase64String(signatureData.SignedDocument.Document);
        //         string outputFilePath = Path.Combine("Signed Documents", $"results_{signatureData.UniqueId}.pdf");

        //         using (var tx = new ServerTextControl())
        //         {
        //             tx.Create();

        //             // Load the document from Base64
        //             tx.Load(signedDocumentBytes, BinaryStreamType.InternalUnicodeFormat);

        //             // Load digital certificate
        //             var certificate = new X509Certificate2(CertificatePath, CertificatePassword, X509KeyStorageFlags.Exportable);

        //             // Assign the certificate to signature field
        //             var saveSettings = new SaveSettings
        //             {
        //                 CreatorApplication = "TX Text Control Blazor Sample Application",
        //                 SignatureFields = new[]
        //                 {
        //                 new DigitalSignature(certificate, null, "txsign")
        //             }
        //             };

        //             // Save as signed PDF
        //             tx.Save(outputFilePath, StreamType.AdobePDF, saveSettings);
        //         }

        //         return Ok(new { message = "Document signed successfully.", filePath = $"Signed Documents/results_{signatureData.UniqueId}.pdf" });
        //     }
        //     catch (Exception ex)
        //     {
        //         // Log the error 
        //         Console.WriteLine($"Error during signing: {ex.Message}");
        //         return StatusCode(500, "An error occurred while signing the document.");
        //     }
        // }
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
        private async Task<string> GetDocumentBytes(string filePath, StreamType streamType)
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
                // tx.Select(200, tx.Text.Length - 200);
                // tx.Clear();

                using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                {
                    // string jsonData = System.IO.File.ReadAllText("Documents/jsonData.json");

                }

                // ConvertToMergeFields(tx);
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
                        // SNOrder dbOrder = GetMappedOrderObj(orderIds[i]);
                        // string jsonData = await GetOrderDataJsonFromDb(orderIds[i], true);

                        // Console.WriteLine(JsonConvert.SerializeObject(dbOrder));
                        tx.Create();

                        var loadSettings = new LoadSettings
                        {
                            ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                            LoadSubTextParts = true
                        };


                        tx.Load("Documents/template_order.docx", StreamType.WordprocessingML, loadSettings);

                        using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                        {
                            mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                            // mailMerge.MergeObject(dbOrder);
                            // mailMerge.MergeJsonData(jsonData);
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

    }
}
