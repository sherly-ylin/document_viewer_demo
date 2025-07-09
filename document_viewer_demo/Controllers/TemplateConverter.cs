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

            doc.SaveAs2(docPath.Replace(".docx", "_updated.docx"));
            doc.Close(false);
            wordApp.Quit();

        }
    }
}