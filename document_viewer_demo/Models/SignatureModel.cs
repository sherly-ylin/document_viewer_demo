using TXTextControl;
using TXTextControl.DocumentServer;
using TXTextControl.Web.MVC.DocumentViewer;
using TXTextControl.Web.MVC.DocumentViewer.Models;

namespace document_viewer_demo.Models
{
  public class SignedDocumentModel
  {
    public SignatureModel SignatureModel { get; set; }
    public Envelope Envelope { get; set; }
    public string SignerId { get; set; }
    public string SignatureImage { get; set; }
  }

  public class SignatureModel
  {
    public string Document { get; set; }
    public int NumPages { get; set; }
    public string SignerInitials { get; set; }
    public string SignerName { get; set; }
    public DateTime TimeStamp { get; set; }
    public string UniqueId { get; set; }
    public string IPAddress { get; set; }

  }
}