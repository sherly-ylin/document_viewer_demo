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

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            try
            {
                // Load template, merge data, and get the merged document
                // string mergedDocumentBase64 = LoadTemplateAndMergeData();
                string mergedDocumentBase64 = LoadTemplateAndMergeMultipleOrders(new List<int> { 7261, 7262, 7264});

                ViewBag.HasDocument = true;
                ViewBag.DocumentData = mergedDocumentBase64;
            }
            catch (Exception ex)
            {
                ViewBag.HasDocument = false;
                ViewBag.ErrorMessage = ex.Message;
                _logger.LogError(ex, "Error processing document template");
            }

            return View();
        }
        private string LoadTemplateAndMergeMultipleOrders(List<int> orderIds)
        {
            Console.WriteLine("Merging multiple orders: " + string.Join(", ", orderIds));
            using (ServerTextControl masterTx = new ServerTextControl())
            {
                masterTx.Create();
                bool isFirstDoc = true;
                foreach (var orderId in orderIds)
                {
                    Console.WriteLine("Processing OrderId: " + orderId);
                    using (ServerTextControl tx = new ServerTextControl())
                    {
                        tx.Create();

                        // Load the template
                        var loadSettings = new LoadSettings
                        {
                            ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                            LoadSubTextParts = true
                        };
                        tx.Load("Documents/template_order.docx", StreamType.WordprocessingML, loadSettings);

                        SNOrder dbOrder = GetOrderFromDb(orderId);

                        using (MailMerge mailMerge = new MailMerge { TextComponent = masterTx })
                        {
                            mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                            mailMerge.MergeObject(dbOrder);
                        }

                        byte[] bytes;
                        tx.Save(out bytes, BinaryStreamType.InternalUnicodeFormat);

                        if (isFirstDoc)
                        {
                            masterTx.Load(bytes, BinaryStreamType.InternalUnicodeFormat);
                            isFirstDoc = false;
                        }
                        else
                        {
                            Console.WriteLine("Appending page break");
                            masterTx.Append("\f", StringStreamType.PlainText, AppendSettings.None);
                            Console.WriteLine("Appending document for OrderId: " + orderId);
                            masterTx.Append(bytes, BinaryStreamType.InternalUnicodeFormat, AppendSettings.None);
                        }
                    }
                }

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
        private string LoadTemplateAndMergeData()
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

                tx.Load("Documents/template_order.docx", StreamType.WordprocessingML, loadSettings);

                // Get data from database
                SNOrder dbOrder = GetOrderFromDb(7262);

                // Merge the data
                using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                {
                    mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                    mailMerge.MergeObject(dbOrder);
                }

                // Save the merged document to a byte array
                byte[] documentBytes;
                var saveSettings = new SaveSettings
                {
                    CreatorApplication = "Document Viewer Demo"
                };

                tx.Save(out documentBytes, BinaryStreamType.InternalUnicodeFormat, saveSettings);

                return Convert.ToBase64String(documentBytes);
            }
        }

        public SNOrder GetOrderFromDb(int orderId)
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
                    var query = @"SELECT * FROM SNOrder o 
                                JOIN SNOrderLine ol on o.OrderId = ol.OrderId
                                WHERE o.OrderId = @OrderId
                                ORDER BY ol.BundleID, Model";
                    var cmd = new Microsoft.Data.SqlClient.SqlCommand(query, conn);

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
                                Model = itemRow["Model"].ToString(),
                                Quantity = Convert.ToInt32(itemRow["Quantity"]),
                                SellPrice = Convert.ToDecimal(itemRow["SellPrice"]),
                                LineTotal = Convert.ToDecimal(itemRow["LineTotal"])
                            });
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

        // Alternative method for getting document as PDF
        public IActionResult GetDocumentAsPdf(int orderId = 7262)
        {
            try
            {
                using (ServerTextControl tx = new ServerTextControl())
                {
                    tx.Create();

                    var loadSettings = new LoadSettings
                    {
                        ApplicationFieldFormat = ApplicationFieldFormat.MSWord,
                        LoadSubTextParts = true
                    };

                    tx.Load("Documents/template_order.docx", StreamType.WordprocessingML, loadSettings);

                    SNOrder dbOrder = GetOrderFromDb(orderId);

                    using (MailMerge mailMerge = new MailMerge { TextComponent = tx })
                    {
                        mailMerge.FormFieldMergeType = FormFieldMergeType.None;
                        mailMerge.MergeObject(dbOrder);
                    }

                    // Export as PDF
                    byte[] pdfBytes;
                    tx.Save(out pdfBytes, BinaryStreamType.AdobePDF);

                    return File(pdfBytes, "application/pdf", $"Order_{orderId}.pdf");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error generating PDF");
                return BadRequest("Error generating PDF: " + ex.Message);
            }
        }

        public IActionResult Privacy()
        {
            return View();
        }
    }
}