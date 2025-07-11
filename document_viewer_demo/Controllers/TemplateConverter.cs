using Word = Microsoft.Office.Interop.Word;
using System.Text.RegularExpressions;

namespace document_viewer_demo.Controllers
{
    public class TemplateConverter
    {
        string docPath;

        public TemplateConverter(string dp)
        {
            docPath = dp;
            Console.WriteLine("File path set");
        }

        public void ConvertMergeFields()
        {
            Console.WriteLine("Entering [ConvertMergeFields]");
            Console.WriteLine(docPath);
            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Open(docPath);
            wordApp.Visible = false;

            var regex = new Regex(@"\{\{(.*?)\}\}");

            foreach (Word.Range range in doc.StoryRanges)
            {
                Word.Range currentRange = range;
                do
                {
                    var matches = regex.Matches(currentRange.Text);
                    foreach (Match match in matches)
                    {
                        string fullTag = match.Groups[0].Value; // e.g. {{OR.OrderID}}
                        string fieldName = match.Groups[1].Value; // e.g. OR.OrderID

                        Word.Find find = currentRange.Find;
                        find.Text = fullTag;
                        find.Replacement.Text = "";
                        find.Forward = true;
                        find.Wrap = Word.WdFindWrap.wdFindStop;

                        if (find.Execute())
                        {
                            Word.Range matchRange = currentRange.Duplicate;
                            matchRange.Text = ""; // delete the tag
                            matchRange.Fields.Add(matchRange, Word.WdFieldType.wdFieldMergeField, fieldName);
                        }
                    }

                    currentRange = currentRange.NextStoryRange;
                }
                while (currentRange != null);
            }

            doc.SaveAs2("C:\\OBD_updated.docx");
            doc.Close(false);
            wordApp.Quit();

        }

        public void ExtractMergeFields()
        {
            var wordApp = new Word.Application();
            var doc = wordApp.Documents.Open(docPath);
            wordApp.Visible = false;

            var regex = new Regex(@"\{\{(.*?)\}\}");
            var fieldNames = new List<string>();

            foreach (Word.Range range in doc.StoryRanges)
            {
                Word.Range currentRange = range;
                Console.WriteLine("======== new word range");
                do
                {
                    var matches = regex.Matches(currentRange.Text);
                    foreach (Match match in matches)
                    {
                        string fieldName = match.Groups[1].Value.Trim();
                        fieldNames.Add(fieldName);
                    }

                    currentRange = currentRange.NextStoryRange;
                }
                while (currentRange != null);
            }

            doc.Close(false);

            // Create new doc to save field names
            var newDoc = wordApp.Documents.Add();
            Word.Paragraph para = newDoc.Content.Paragraphs.Add();

            foreach (string name in fieldNames)
            {
                Console.WriteLine(name);

                para.Range.InsertBefore(name + "\n");
            }

            newDoc.SaveAs2(docPath.Replace(".docx", "_fields.docx"));
            newDoc.Close();
            wordApp.Quit();
        }

        public void ConvertQueryJson(string filePath)
        {
            string data = System.IO.File.ReadAllText(filePath);
            var beginRegex = new Regex(@"\+\s\'\<(OR.*?)\>\'\s\+");
            var beginRegex2 = new Regex(@" \'\<(OR.*?)\>\'\s*\+");
            var closeRegex = new Regex(@"\+\s\'\<\/(OR\..*?)\>\'");
            var closeRegex2 = new Regex(@"\+\s\'\<\/(OR.*?)\>\'\s*\+");
            var ORRegex = new Regex(@"(OR\.)");
            var PlusRegex = new Regex(@"\+\s");
            var formatRegex = new Regex(@"(dbo\.fn_FormatCurrency\((.*?)\))");
            var formatRegex2 = new Regex(@"(dbo\.fn_FormatXMLChars\((.*?)\))");
            var formatRegex3 = new Regex(@"(dbo\.fn_formatnumber\((.*?)\))");
            // Remove any text that matches beginRegex
            data = beginRegex2.Replace(data, "");

            // Replace any text that matches closeRegex with the text inside the tag
            // data = formatRegex.Replace(data, m => m.Groups[1].Value);
            // data = formatRegex2.Replace(data, m => m.Groups[1].Value);
            // data = formatRegex3.Replace(data, m => m.Groups[1].Value);
            data = closeRegex2.Replace(data, m => $"AS {m.Groups[1].Value}, \n\t");
            // data = ORRegex.Replace(data, "AS ");

            Console.WriteLine(data);

            string outputPath = filePath.Replace(".txt", "_processed.txt");
            System.IO.File.WriteAllText(outputPath, data);
            Console.WriteLine($"Processed data saved to {outputPath}");
        }
    }
}