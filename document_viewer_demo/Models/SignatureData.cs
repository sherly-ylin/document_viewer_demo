using TXTextControl;
using TXTextControl.DocumentServer;
using TXTextControl.Web.MVC.DocumentViewer;
using TXTextControl.Web.MVC.DocumentViewer.Models;

public class SignatureData {
  public SignatureDocument SignedDocument { get; set; }
  public string DocumentName { get; set; }
  public string SignatureImage { get; set; }
  public string Name { get; set; }
  public string Initials { get; set; }
  public double InitialsWidth { get; set; }
  public string UniqueId { get; set; }
  public string SignatureBoxName { get; set; }
  public SignatureBox[] SignatureBoxes { get; set; }
  public string DocumentData { get; set; }
  public object CustomSignatureData { get; set; }
  public List<CompletedFormField> FormFields { get; set; } = new List<CompletedFormField>();
  public string RedirectUrl { get; set; }
  public bool CustomSigning { get; set; }
}