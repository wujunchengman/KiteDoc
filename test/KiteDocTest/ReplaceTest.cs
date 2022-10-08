using DocumentFormat.OpenXml.Packaging;
using KiteDoc.Interface;
using KiteDoc;
using System;
using Xunit;
using System.IO;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System.Linq;
using System.Collections.Generic;
using System.Diagnostics;
using Xunit.Abstractions;

namespace KiteDocTest
{
    public class ReplaceTest
    {
        private readonly ITestOutputHelper output;

        public ReplaceTest(ITestOutputHelper output)
        {
            this.output = output;
        }

        [Fact]
        public void ReplaceText()
        {
            var filename = "�滻�ı�.docx";
            string testPath = CopyTestFile(filename);




            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var count = doc.Replace("��װ", "�ı�֮���Ŀ���ı�");
                Assert.Equal(4, count);
            }
            //File.Delete(testPath);
        }

        public Table TestTable { get {
                #region �������
                var data = new string[3, 3];

                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        data[i, j] = "����" + i + j;
                    }
                }

                Table table = new Table();

                TableProperties props = new TableProperties(
                    new TableBorders(
                    new TopBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new BottomBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new LeftBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new RightBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideHorizontalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    },
                    new InsideVerticalBorder
                    {
                        Val = new EnumValue<BorderValues>(BorderValues.Single),
                        Size = 12
                    }));

                table.AppendChild<TableProperties>(props);

                for (var i = 0; i <= data.GetUpperBound(0); i++)
                {
                    var tr = new TableRow();
                    for (var j = 0; j <= data.GetUpperBound(1); j++)
                    {
                        var tc = new TableCell();
                        tc.Append(new Paragraph(new Run(new Text(data[i, j]))));

                        // Assume you want columns that are automatically sized.
                        tc.Append(new TableCellProperties(
                            new TableCellWidth { Type = TableWidthUnitValues.Auto }));

                        tr.Append(tc);
                    }
                    table.Append(tr);
                }

                #endregion
                return table;
            }
        }

        [Fact]
        public void ReplaceTextToTable()
        {
            var filename = "�滻�ı�Ϊ���.docx";
            string testPath = CopyTestFile(filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("���", table);
                Assert.Equal(3, count);
            }
        }

        [Fact]
        public void ReplaceTableTextToNestTable()
        {
            var filename = "�滻����е��ı��γ�Ƕ�ױ��.docx";
            string testPath = CopyTestFile(filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("���", table);
                Assert.Equal(1, count);
            }

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, false))
            {
                Assert.Equal(2, doc.MainDocumentPart.Document.Descendants<Table>().Count());
            }
        }

        [Fact]
        public void ReplaceTextToTableSaveFormatting()
        {
            var filename = "�滻�ı�Ϊ���.docx";
            string testPath = CopyTestFile(filename);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("���", table,true);
                Assert.Equal(3, count);
            }
            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, false))
            {

                var c = doc.MainDocumentPart.Document.Descendants<Table>().ToArray();

                

                var runs = c[1].Descendants<Run>();
                
                var dest = runs.Select(x => x.RunProperties.FontSize.Val.Value);
                foreach (var item in dest)
                {
                    output.WriteLine(item);
                }
                Assert.Contains(40.ToString(),dest );

            }
        }

        [Fact]
        public void ReplaceTextToTableSplitRun()
        {
            var filename = "�滻�ֶ�Run���ı�.docx";
            string testPath = CopyTestFile(filename);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("{GroupMembers}", table, true);
                Assert.Equal(2, count);
            }
        }

        private static string CopyTestFile(string filename)
        {
            var filePath = "StaticResource" + Path.DirectorySeparatorChar + filename;
            var testPath = "StaticResource" + Path.DirectorySeparatorChar + "test" + filename;
            File.Copy(filePath, testPath, true);
            return testPath;
        }
    }
}
