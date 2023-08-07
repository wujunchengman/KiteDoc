using DocumentFormat.OpenXml.Packaging;
using Xunit;
using KiteDoc;
using KiteDoc.ElementBuilder;
using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Drawing.Charts;

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

        [Fact]
        public void DataSplitParagraphTest()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(testPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                var table = new TableBuilder();
                var data = new List<List<string?>>()
                {
                    new List<string?>{"AAAAAAAAAAAAAAVAAAAAA","BBBBBBBBBBVBBBBBBBB"},
                    new List<string?>{"AAAAAAAAAAAAAAVAAAAAA","BBBBBBBBBBVBBBBBBBB"},
                };
                table
                    .SetDataSplitParagraph("V",false)
                    .SetTableData(data).SetHorizationNullMerge().SetBorder().SetJustification();

                body.AppendChild(table.Build());


            }
        }

        [Fact]
        public void DataNullFirstTest()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(testPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                var table = new TableBuilder();
                var data = new List<List<string?>>()
                {
                    new List<string?>{"测试",null,null},
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                };

                var width = new float[] { 33.3f, 33.3f, 33.4f };

                table
                    .SetTableData(data)
                    .SetBorder()
                    .SetJustification()
                    .SetHorizationNullMerge()
                    .SetDataSplitParagraph("#$$#", false)
                    .SetTableCellWidth(KiteDoc.ElementBuilder.TableWidthType.Percent, width)
                    .SetTableDataFontSize(10.5f);

                body.AppendChild(table.Build());


            }
        }
        
        [Fact]
        public void DataNullFirstNotSplitParagraphTest()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(testPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                var table = new TableBuilder();
                var data = new List<List<string?>>()
                {
                    new List<string?>{"测试",null,null},
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                };

                var width = new float[] { 33.3f, 33.3f, 33.4f };

                table
                    .SetTableData(data)
                    .SetBorder()
                    .SetJustification()
                    .SetHorizationNullMerge()
                    .SetTableCellWidth(KiteDoc.ElementBuilder.TableWidthType.Percent, width)
                    .SetTableDataFontSize(10.5f);

                body.AppendChild(table.Build());


            }
        }

        [Fact]
        public void DataNoteNullTest()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(testPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                var table = new TableBuilder();
                var data = new List<List<string?>>()
                {
                    new List<string?>{"测试","测试","测试"},
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                };

                var width = new float[] { 33.3f, 33.3f, 33.4f };

                table
                    .SetTableData(data)
                    .SetBorder()
                    .SetJustification()
                    .SetHorizationNullMerge()
                    .SetDataSplitParagraph("#$$#", false)
                    .SetTableCellWidth(KiteDoc.ElementBuilder.TableWidthType.Percent, width)
                    .SetTableDataFontSize(10.5f);

                body.AppendChild(table.Build());


            }
        }

        [Fact]
        public void DataFontSizeTest()
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Create(testPath, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
            {

                // Add a main document part. 
                MainDocumentPart mainPart = doc.AddMainDocumentPart();

                // Create the document structure and add some text.
                mainPart.Document = new Document();
                Body body = mainPart.Document.AppendChild(new Body());

                var table = new TableBuilder();
                var data = new List<List<string?>>()
                {
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                    new List<string?>{"AAA","AAAA","AAAAAAA"},
                };

                var width = new float[] { 33.3f, 33.3f, 33.4f };

                table
                    .SetTableHeader(new List<string> { "AAAA", "BBBBBB","CCCCCC"})
                    .SetTableData(data)
                    .SetBorder()
                    .SetJustification()
                    .SetHorizationNullMerge()
                    .SetDataSplitParagraph("#$$#", false)
                    .SetTableHeaderCellWidth(KiteDoc.ElementBuilder.TableWidthType.Percent, width)
                    .SetTableDataFontSize(21);

                body.AppendChild(table.Build());


            }
        }
    }
}
