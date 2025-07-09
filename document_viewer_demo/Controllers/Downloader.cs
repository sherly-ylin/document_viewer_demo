using TXTextControl;

namespace document_viewer_demo.Controllers
{
    class Downloader
    {
        byte[] documentBytes { get; set; }

        public Downloader(byte[] bytes)
        {
            documentBytes = bytes;
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

    }
}