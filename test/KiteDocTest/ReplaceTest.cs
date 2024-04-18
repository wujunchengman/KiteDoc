using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc;
using KiteDoc.ElementBuilder;
using KiteDocTest.Utils;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using Xunit;
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
            var filename = "替换文本.docx";
            string testPath = FileUtils.CopyTestFile(filename, "ReplaceText" + filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var count = doc.Replace("假装", "改变之后的目标文本");
                Assert.Equal(5, count);
            }
            //File.Delete(testPath);
        }

        [Fact]
        public void ReplaceText_2_Test()
        {
            var filename = "保温箱验证方案.docx";
            string testPath = FileUtils.CopyTestFile(filename, "ReplaceText_2_Test" + filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var count = doc.Replace("{DescriptionA}", "保温箱");
                Assert.Equal(2, count);
            }
            //File.Delete(testPath);
        }

        [Fact]
        public void ReplaceTextSplitRun()
        {
            var filename = "替换分段Run的文本.docx";
            string testPath = FileUtils.CopyTestFile(filename, "ReplaceTextSplitRun" + filename);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var count = doc.Replace("{GroupMembers}", "新的字符串");
                Assert.Equal(2, count);
            }
        }

        public Table TestTable
        {
            get
            {
                #region 构建表格

                var data = new string[3, 3];

                for (int i = 0; i < 3; i++)
                {
                    for (int j = 0; j < 3; j++)
                    {
                        data[i, j] = "测试" + i + j;
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
                            new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth { Type = TableWidthUnitValues.Auto }));

                        tr.Append(tc);
                    }
                    table.Append(tr);
                }

                #endregion 构建表格

                return table;
            }
        }

        [Fact]
        public void ReplaceTextToTable()
        {
            var filename = "替换文本为表格.docx";
            string testPath = FileUtils.CopyTestFile(filename, "ReplaceTextToTable" + filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("表格", table);
                Assert.Equal(3, count);
            }
        }

        [Fact]
        public void ReplaceTableTextToNestTable()
        {
            var filename = "替换表格中的文本形成嵌套表格.docx";
            string testPath = FileUtils.CopyTestFile(filename, "ReplaceTableTextToNestTable" + filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("表格", table);
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
            var filename = "替换文本为表格.docx";
            string testPath = FileUtils.CopyTestFile(filename, "ReplaceTextToTableSaveFormatting" + filename);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("表格", table, true);
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
                Assert.Contains(40.ToString(), dest);
            }
        }

        [Fact]
        public void ReplaceTextToTableSplitRun()
        {
            var filename = "替换分段Run的文本.docx";
            string testPath = FileUtils.CopyTestFile(filename, "ReplaceTextToTableSplitRun" + filename);
            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var table = TestTable;
                var count = doc.Replace("{GroupMembers}", table, true);
                Assert.Equal(2, count);
            }
        }
    }
}