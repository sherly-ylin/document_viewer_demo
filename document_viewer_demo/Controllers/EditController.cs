using document_viewer_demo.Models;
using Microsoft.AspNetCore.Mvc;
using TXTextControl;
using TXTextControl.DocumentServer;

namespace document_viewer_demo.Controllers
{
    public class EditController : Controller
    {
        private readonly ILogger<EditController> _logger;

        public EditController(ILogger<EditController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            try
            {
                _logger.LogInformation("Document not found in session, generating new document");
                string docBase64 = LoadDocument("Documents/edit.tx", StreamType.InternalFormat);

                ViewBag.HasDocument = true;
                ViewBag.DocumentData = docBase64;
                ViewBag.DocumentName = $"Edit_{DateTime.Now:yyyyMMdd_HHmmss}.docx";
            }
            catch (Exception ex)
            {
                ViewBag.HasDocument = false;
                ViewBag.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Error processing document template");
            }

            return View();
        }
        
        public string LoadDocument(string filePath, StreamType streamType)
        {
            byte[] document;
            // Load the document from the specified path and return it as a Base64 string
            using (var tx = new TXTextControl.ServerTextControl())
            {
                tx.Create();
                tx.Load(filePath, streamType);
                {
                    tx.Save(out document, BinaryStreamType.InternalUnicodeFormat);
                    return Convert.ToBase64String(document);
                }
            }
        }
    }
}