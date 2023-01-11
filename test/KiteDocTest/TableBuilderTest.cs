using DocumentFormat.OpenXml.Packaging;
using Xunit;
using KiteDoc;
using KiteDoc.ElementBuilder;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;

namespace KiteDocTest
{
    public class TableBuilderTest
    {

        private string testPath = "test.docx";

        [Fact]
        public void BuildMergeTable()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(testPath,DocumentFormat.OpenXml.WordprocessingDocumentType.Document,true))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                var table = new TableBuilder();
                var data = new List<List<string?>>();
                data.Add(new List<string?> { "a", "b", "c", "d","e" });
                data.Add(new List<string?> { "A", "B", "C", null, "E" });
                data.Add(new List<string?> { "A", "B", "C", "D", null }); 
                data.Add(new List<string?> { "A", "B", null, null, "E" });
                table.SetTableData(data).SetHorizationNullMerge().SetBorder().SetJustification();

                body.AppendChild( table.Build());


            }
        }
    }
}
