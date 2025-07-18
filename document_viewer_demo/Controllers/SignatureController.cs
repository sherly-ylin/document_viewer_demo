using document_viewer_demo.Models;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using TXTextControl;
using TXTextControl.DocumentServer;
using TXTextControl.Web.MVC.DocumentViewer.Models;
using System.Security.Cryptography.X509Certificates;

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

        [HttpPost]
        public IActionResult HandleSignature([FromBody] SignatureData data)
        {
            try
            {
                Console.WriteLine("=== HandleSignature called ===");
                byte[] pdfBytes;

                using (var tx = new TXTextControl.ServerTextControl())
                {
                    tx.Create();
                    tx.Load(Convert.FromBase64String(data.SignedDocument.Document), BinaryStreamType.InternalUnicodeFormat);
                    byte[] signatureImage = Convert.FromBase64String(data.SignedDocument.SignatureBoxMergeResults[0].ImageResult);

                    X509Certificate2 cert = new X509Certificate2("App_Data/testesigncert.pfx", "test123");
                    var timeStampServer = "http://timestamp.digicert.com";

                    List<DigitalSignature> signatures = new List<DigitalSignature>();

                    foreach (SignatureField field in tx.SignatureFields)
                    {
                        // field.Name = Guid.NewGuid().ToString();
                        Console.WriteLine("=== Processing Signature Field: " + field.Name + " ===");
                        signatures.Add(new DigitalSignature(null, null, field.Name));
                    }

                    SaveSettings saveSettings = new SaveSettings()
                    {
                        CreatorApplication = "testesign",
                        SignatureFields = signatures.ToArray()
                    };

                    tx.Save(out pdfBytes, BinaryStreamType.AdobePDFA, saveSettings);
                    tx.Save($"App_Data/signed_{DateTime.Now:yyyyMMdd_HHmmss}.pdf", StreamType.AdobePDFA, saveSettings);
                }
                Console.WriteLine("=== PDF bytes generated ===");
                return File(pdfBytes, "application/pdf", "SignedDoc.pdf");
                // return Ok(new { message = "Document signed successfully.", filePath = $"Signed Documents/results_{signatureData.UniqueId}.pdf" });
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, $"Error handling signature");
                return StatusCode(StatusCodes.Status500InternalServerError, $"An error occurred while processing the signature. {ex.Message}");
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
