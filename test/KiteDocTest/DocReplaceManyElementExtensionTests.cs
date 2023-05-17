using KiteDoc;
using KiteDocTest.Utils;
using System.Collections.Generic;
using Xunit;
using Xunit.Abstractions;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Packaging;
using KiteDoc.ElementBuilder;
using DocumentFormat.OpenXml;

#nullable enable

namespace KiteDoc.Tests
{
    public class DocReplaceManyElementExtensionTests
    {
        private readonly ITestOutputHelper _output;

        public DocReplaceManyElementExtensionTests(ITestOutputHelper output)
        {
            _output = output;
        }

        [Fact]
        public void ReplaceTextToParagraphs()
        {
            var filename = "替换文本.docx";
            string testPath = FileUtils.CopyTestFile(filename);




            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var builder = new ParagraphBuilder();
                builder.AppendText("测试文本");

                var p = builder.Build();

                var ps = new List<Paragraph>();

                for (int i = 0; i < 5; i++)
                {
                    ps.Add(p.Clone() as Paragraph);
                }


                var count = doc.Replace("假装", ps);
                Assert.Equal(5, count);
            }
            //File.Delete(testPath);
        }

        [Fact]
        public void ReplaceStringToTablesTest()
        {
                var filename = "替换文本为多个表格.docx";
                string testPath = FileUtils.CopyTestFile(filename);




                using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
                {
                    var builder = new TableBuilder();
                    builder.SetTableData(new List<List<string?>>
                    {
                        new List<string?>{"AAAA","BBBB","CCCC"},
                        new List<string?>{"AAAA","BBBB","CCCC"},
                        new List<string?>{"AAAA","BBBB","CCCC"},
                        new List<string?>{"AAAA","BBBB","CCCC"},
                    }).SetTableCellWidth(ElementBuilder.TableWidthType.Percent,new float[] {30,30,40})
                    .SetBorder()
                    .SetJustification()
                    ;

                    var t = builder.Build();



                var ts = new List<Table>();

                    for (int i = 0; i < 5; i++)
                    {
                        ts.Add((Table)t.Clone());
                    }


                    var count = doc.Replace("替换", ts,true);
                    Assert.Equal(1, count);
                }
        }


        [Fact]
        public void ReplaceStringToTablesAndParagraphTest()
        {
            var filename = "替换文本为多个表格.docx";
            string testPath = FileUtils.CopyTestFile(filename);


            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var builder = new TableBuilder();
                builder.SetTableData(new List<List<string?>>
                    {
                        new List<string?>{"AAAA","BBBB","CCCC"},
                        new List<string?>{"AAAA","BBBB","CCCC"},
                        new List<string?>{"AAAA","BBBB","CCCC"},
                        new List<string?>{"AAAA","BBBB","CCCC"},
                    }).SetTableCellWidth(ElementBuilder.TableWidthType.Percent, new float[] { 30, 30, 40 })
                .SetBorder()
                .SetJustification()
                ;

                var t = builder.Build();

                var ts = new List<OpenXmlCompositeElement>();

                var pBuilder = new ParagraphBuilder();
                pBuilder.AppendText("测试文本");


                var p = pBuilder.Build();

                ts.Add((Paragraph)p.Clone());

                for (int i = 0; i < 5; i++)
                {
                    ts.Add((Table)t.Clone());
                }

                ts.Add((Paragraph)p.Clone());

                var count = doc.Replace("替换", ts);
                Assert.Equal(1, count);
            }
        }
    }
}