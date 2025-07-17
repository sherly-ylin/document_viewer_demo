using document_viewer_demo.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using TXTextControl;
using TXTextControl.DocumentServer;
using TXTextControl.Web.MVC.DocumentViewer.Models;

namespace document_viewer_demo.Controllers
{
    public class SignatureController : Controller
    {
        private readonly ILogger<SignatureController> _logger;

        string connectionString = "Server=192.168.20.97;Database=SalesChain0602_MS_MN;User Id=ylin;Password=9244@Wahg;TrustServerCertificate=True;";


        // private List<int> pageLengths { get; set; } = new List<int>();
        public SignatureController(ILogger<SignatureController> logger)
        {
            _logger = logger;
        }

        public Task<IActionResult> Index()
        {
            try
            {
                _logger.LogInformation("Document not found in session, generating new document");
                string docBase64 = "";

                docBase64 = LoadDocument("Documents/signature.tx", StreamType.InternalFormat);

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

            return Task.FromResult<IActionResult>(View());
        }

        public IActionResult Sign(string id)
        {
            try
            {

                string document = LoadDocument("Documents/signature.tx", StreamType.InternalFormat);
                // Envelope envelope = new Envelope() {
                //     EnvelopeID = id, 
                //     UserID = "testUser",
                //     Sender = "Test Sender",
                //     Name = "Test Envelope",
                //     Created = DateTime.Now,
                //     Sent = DateTime.Now.AddMinutes(5),
                //     Status = EnvelopeStatus.Incomplete,
                //     ContainsSignatureBoxes = true,
                //     SignatureInformation = new SignatureModel() {
                //         Document = document,
                //         NumPages = 1,
                //         SignerInitials = "TS",
                //         SignerName = "Test Signer",
                //         TimeStamp = DateTime.Now,
                //         UniqueId = Guid.NewGuid().ToString(),
                //         IPAddress = HttpContext.Connection.RemoteIpAddress?.ToString() ?? "Unknown"
                //     }
                // };

                SignModel model = new SignModel()
                {
                    Document = document
                    // Envelope = envelope,
                    // Signer = currentSigner
                };

                return View(model);
            }
            catch (Exception ex)
            {
                ViewBag.HasDocument = false;
                ViewBag.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Error processing document template");
                return View("Error", new { message = "An error occurred while processing the document." });
            }

        }

        private string LoadDocument(string filePath, StreamType streamType)
        {
            using (ServerTextControl tx = new ServerTextControl())
            {
                tx.Create();

                tx.Load(filePath, streamType);

                using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                {
                    // string jsonData = System.IO.File.ReadAllText("Documents/jsonData.json");

                }
                byte[] bytes;
                tx.Save(out bytes, BinaryStreamType.InternalUnicodeFormat);
                return Convert.ToBase64String(bytes);
            }
        }

        [HttpPost]
        public IActionResult HandleSignature([FromBody] SignatureData data)
        {
            byte[] pdfBytes;

            using (var tx = new TXTextControl.ServerTextControl())
            {
                tx.Create();
                tx.Load(Convert.FromBase64String(data.SignedDocument.Document), BinaryStreamType.InternalUnicodeFormat);
                byte[] signatureImage = Convert.FromBase64String(data.SignedDocument.SignatureBoxMergeResults[0].ImageResult);
                // Save without digital cert
                tx.Save(out pdfBytes, TXTextControl.BinaryStreamType.AdobePDFA);
            }

            return File(pdfBytes, "application/pdf", "SignedDoc.pdf");
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
    }
}
