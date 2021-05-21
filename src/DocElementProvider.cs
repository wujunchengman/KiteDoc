using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.Enum;
using KiteDoc.Interface;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using A = DocumentFormat.OpenXml.Drawing;
using DW = DocumentFormat.OpenXml.Drawing.Wordprocessing;
using PIC = DocumentFormat.OpenXml.Drawing.Pictures;

namespace KiteDoc
{
    public class DocElementProvider : IDocElementProvider
    {
        /// <summary>
        /// 返回一个BookmarkStart对象，每一个BookmarkStart必须要有Id相同的BookmarkEnd与之对应
        /// </summary>
        /// <param name="bookmarkId">Bookmark(书签)的唯一标识，用于绑定BookmarkEnd(书签结束)</param>
        /// <param name="bookmarkName">书签名</param>
        /// <returns></returns>
        public BookmarkStart GetBookmarkStartObject(int bookmarkId, string bookmarkName)
        {
            BookmarkStart bookmarkStart = new BookmarkStart();
            bookmarkStart.Id = bookmarkId.ToString();
            bookmarkStart.Name = bookmarkName;

            return bookmarkStart;
        }

        /// <summary>
        /// 返回一个BookmarkEnd对象，其Id必须与BookmarkStart的Id相同才能表示一个完整书签
        /// </summary>
        /// <param name="bookmarkId">Bookmark(书签)的唯一标识，必须与BookmarkStart的Id相同才能表示一个完整书签</param>
        /// <returns></returns>
        public BookmarkEnd GetBookmarkEndObject(int bookmarkId)
        {
            BookmarkEnd bookmarkEnd = new BookmarkEnd();
            bookmarkEnd.Id = bookmarkId.ToString();

            return bookmarkEnd;
        }


        /// <summary>
        /// 返回一个Run对象
        /// </summary>
        /// <param name="isBold">是否加粗</param>
        /// <param name="fontSize">字体大小，传递为Null则使用默认字体大小</param>
        /// <param name="font">字体，默认为宋体</param>
        /// <returns></returns>
        public Run GetRunObject(bool isBold, int? fontSize = null, string font = "宋体")
        {
            var run = new Run();
            var runProperties = new RunProperties();
            if (isBold)
            {
                runProperties.AppendChild(new Bold());
            }

            if (fontSize != null)
            {
                runProperties.AppendChild(new FontSize() { Val = (fontSize * 2).ToString() });
            }

            // 统一字体，将Ascii(拼音数字)，HighAnsi，复杂文种，中文设置为同一字体
            runProperties.AppendChild(new RunFonts()
            { Ascii = font, HighAnsi = font, ComplexScript = font, EastAsia = font });

            run.AppendChild(runProperties);
            return run;
        }

        /// <summary>
        /// 返回一个Paragraph段落对象
        /// </summary>
        /// <param name="align">对齐方式，默认为左对齐</param>
        /// <param name="indentationForTheFirstLine">首行缩进，默认为否</param>
        /// <returns></returns>
        public Paragraph GetParagraphObject(
            HorizontalAlign align = HorizontalAlign.Left,
            bool indentationForTheFirstLine = false)
        {
            var paragraph = new Paragraph();
            var paragraphProperties = new ParagraphProperties();

            // 设置段落对齐方式
            switch (align)
            {
                case HorizontalAlign.Left:
                    paragraphProperties.Justification = new Justification()
                    { Val = new EnumValue<JustificationValues>(JustificationValues.Left) };
                    break;
                case HorizontalAlign.Center:
                    paragraphProperties.Justification = new Justification()
                    { Val = new EnumValue<JustificationValues>(JustificationValues.Center) };
                    break;
                case HorizontalAlign.Right:
                    paragraphProperties.Justification = new Justification()
                    { Val = new EnumValue<JustificationValues>(JustificationValues.Right) };
                    break;
            }

            if (indentationForTheFirstLine)
            {
                paragraphProperties.AppendChild(new Indentation() { FirstLineChars = 200 });
            }

            paragraph.AppendChild(paragraphProperties);
            return paragraph;
        }

        /// <summary>
        /// 返回一个单元格对象
        /// </summary>
        /// <param name="width">宽度百分比</param>
        /// <param name="verticalAlign">垂直对齐方式，默认居中对齐</param>
        /// <param name="cellMerge">单元格合并，默认为普通，即不参与合并</param>
        /// <returns></returns>
        public TableCell GetTableCellObject(
            int width,
            VerticalAlign verticalAlign = VerticalAlign.Center,
            DocCellMergeEnum cellMerge = DocCellMergeEnum.Normal
        )
        {
            TableCell tableCell = new TableCell();
            var tableCellProperties = new TableCellProperties();

            // 指定表格单元格宽度
            tableCellProperties.AppendChild(new TableCellWidth()
            { Type = TableWidthUnitValues.Pct, Width = (width * 50).ToString() });

            // 指定表格单元格
            switch (verticalAlign)
            {
                case VerticalAlign.Center:
                    tableCellProperties.AppendChild(
                        new TableCellVerticalAlignment()
                        { Val = new EnumValue<TableVerticalAlignmentValues>(TableVerticalAlignmentValues.Center) });
                    break;
                case VerticalAlign.Top:
                    tableCellProperties.AppendChild(new TableCellVerticalAlignment()
                    { Val = new EnumValue<TableVerticalAlignmentValues>(TableVerticalAlignmentValues.Top) });
                    break;
                case VerticalAlign.Bottom:
                    tableCellProperties.AppendChild(new TableCellVerticalAlignment()
                    { Val = new EnumValue<TableVerticalAlignmentValues>(TableVerticalAlignmentValues.Bottom) });
                    break;
            }


            // 添加合并单元格标记
            switch (cellMerge)
            {
                case DocCellMergeEnum.Normal:
                    break;
                case DocCellMergeEnum.HorizontalStart:
                    Console.WriteLine("start");
                    tableCellProperties.AppendChild(new HorizontalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    break;
                case DocCellMergeEnum.HorizontalContinue:
                    Console.WriteLine("continue");
                    tableCellProperties.AppendChild(new HorizontalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Continue) });
                    break;
                case DocCellMergeEnum.VerticalStart:
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    break;
                case DocCellMergeEnum.VerticalContinue:
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Continue) });
                    break;
                case DocCellMergeEnum.HorizontalAndVerticalStart:
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    break;
            }

            tableCell.AppendChild(tableCellProperties);

            return tableCell;
        }

        /// <summary>
        /// 返回一个表格对象
        /// </summary>
        /// <param name="isBorder">是否启用表格边框，默认启用</param>
        /// <param name="width">表格宽度(百分比),默认100</param>
        /// <returns></returns>
        public Table GetTableObject(bool isBorder = true, int width = 100)
        {
            Table table = new Table();
            TableProperties tableProperties = new TableProperties();
            tableProperties.AppendChild(new TableWidth()
            { Type = TableWidthUnitValues.Pct, Width = (width * 50).ToString() });

            // 是否启用边框
            if (isBorder)
            {
                // 单框线，Size=2 为1磅粗
                tableProperties.AppendChild(new TableBorders(new TopBorder()
                { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new BottomBorder()

                    { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new LeftBorder()
                    { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new RightBorder()
                    { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new InsideHorizontalBorder()
                    { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 },
                    new InsideVerticalBorder()
                    { Val = new EnumValue<BorderValues>(BorderValues.Single), Size = 2 }));
            }

            table.AppendChild(tableProperties);
            return table;
        }

        /// <summary>
        /// 生成表格，如果单元格没数据会自动合并单元格
        /// </summary>
        /// <param name="tableHead">表头，有空行添加空数据</param>
        /// <param name="data">表格数据</param>
        /// <param name="widthList">表格宽度列表（单位：%）</param>
        /// <param name="fontSize">字号，为null时使用默认字号</param>
        /// <param name="align">对齐方式，默认居中对齐</param>
        /// <param name="isSerialNumber">单元格内换行数据是否编号</param>
        /// <returns></returns>
        public Table GenerateNormalTable(List<string> tableHead, List<string[]> data, List<int> widthList,
            int? fontSize = null, HorizontalAlign align = HorizontalAlign.Center, bool isSerialNumber = false)
        {
            var table = GetTableObject(true, 100);

            if ((tableHead != null) && tableHead.Count != 0)
            {
                TableRow tableRow = new TableRow();
                for (int i = 0; i < tableHead.Count; i++)
                {
                    var tableCell = GetTableCellObject(widthList[i]);

                    var paragraph = GetParagraphObject(HorizontalAlign.Center);

                    var run = GetRunObject(true, fontSize);

                    // 将标题内容添加到run对象中
                    run.AppendChild(new Text(tableHead[i]));
                    // 将run对象添加到段落中
                    paragraph.AppendChild(run);
                    // 将段落对象添加到单元格中
                    tableCell.AppendChild(paragraph);
                    // 将单元格添加到行中
                    tableRow.AppendChild(tableCell);
                }

                // 将表头行添加到表格对象中
                table.AppendChild(tableRow);
            }

            var rowLength = data[0].Length;
            foreach (string[] rowList in data)
            {
                TableRow tableRow = new TableRow();
                for (int i = 0; i < rowLength; i++)
                {
                    // 单元格对象
                    TableCell tableCell = null;
                    // 如果下一单元格没有数据，则开启合并单元格

                    if (string.IsNullOrEmpty(rowList[i]))
                    {
                        tableCell = GetTableCellObject(widthList[i],
                            cellMerge: DocCellMergeEnum.HorizontalContinue);
                    }

                    else if ((i + 1) < rowLength && string.IsNullOrEmpty(rowList[i + 1]))
                    {
                        Console.WriteLine(rowList[i]);
                        tableCell = GetTableCellObject(widthList[i], VerticalAlign.Center,
                            DocCellMergeEnum.HorizontalStart);
                    }
                    else
                    {
                        tableCell = GetTableCellObject(widthList[i], VerticalAlign.Center);
                    }

                    rowList[i] = rowList[i] ?? "";
                    var text = rowList[i].Split("#$$#");

                    if (text.Length > 1)
                    {
                        for (int j = 0; j < text.Length; j++)
                        {
                            var run = GetRunObject(false, fontSize);

                            var paragraph = GetParagraphObject(align);

                            run.AppendChild(isSerialNumber ? new Text((j + 1) + "）" + text[j]) : new Text(text[j]));

                            // 将run对象写入段落
                            paragraph.AppendChild(run);
                            // 将段落写入单元格
                            tableCell.AppendChild(paragraph);
                        }
                    }
                    else
                    {
                        foreach (string s in text)
                        {
                            var run = GetRunObject(false, fontSize);

                            var paragraph = GetParagraphObject(align);

                            // 向Run对象中写入数据
                            run.AppendChild(new Text(s));
                            // 将run对象写入段落
                            paragraph.AppendChild(run);
                            // 将段落写入单元格
                            tableCell.AppendChild(paragraph);
                        }
                    }


                    //将单元格写入行
                    tableRow.AppendChild(tableCell);
                }

                // 将行写入单元格
                table.AppendChild(tableRow);
            }

            return table;
        }

        /// <summary>
        /// 生成表格，如果单元格没数据会自动合并单元格
        /// 使用 #$$# 在单元格内换行
        /// </summary>
        /// <param name="tableHead">表头，有空行添加空数据</param>
        /// <param name="data">表格数据</param>
        /// <param name="widthList">表格宽度列表（单位：%）</param>
        /// <param name="fontSize">字体大小，默认为null，为null时继承上一级字体大小</param>
        /// <param name="align">全局对齐方式，默认居中对齐</param>
        /// <param name="isSerialNumber">单元格内换行数据是否编号</param>
        /// <returns></returns>
        public Table GenerateNormalTable(List<string> tableHead, List<string[]> data, List<int> widthList,
            int? fontSize = null, List<HorizontalAlign> align = null, bool isSerialNumber = false)
        {
            var table = GetTableObject(true, 100);


            if ((tableHead != null) && tableHead.Count != 0)
            {
                TableRow tableRow = new TableRow();
                for (int i = 0; i < tableHead.Count; i++)
                {
                    var tableCell = GetTableCellObject(widthList[i]);

                    var paragraph = GetParagraphObject(HorizontalAlign.Center);

                    var run = GetRunObject(true, fontSize);

                    // 将标题内容添加到run对象中
                    run.AppendChild(new Text(tableHead[i]));
                    // 将run对象添加到段落中
                    paragraph.AppendChild(run);
                    // 将段落对象添加到单元格中
                    tableCell.AppendChild(paragraph);
                    // 将单元格添加到行中
                    tableRow.AppendChild(tableCell);
                }

                // 将表头行添加到表格对象中
                table.AppendChild(tableRow);
            }

            if (align != null && align.Count != 0)
            {
                if (align.Count == data[0].Length)
                {
                    foreach (string[] rowList in data)
                    {
                        TableRow tableRow = new TableRow();
                        for (int i = 0; i < rowList.Length; i++)
                        {
                            TableCell tableCell = null;

                            if (string.IsNullOrEmpty(rowList[i]))
                            {
                                tableCell = GetTableCellObject(widthList[i],
                                    cellMerge: DocCellMergeEnum.HorizontalContinue);
                            }
                            // 如果下一单元格没有数据，则开启合并单元格
                            else if ((i + 1) < rowList.Length && string.IsNullOrEmpty(rowList[i + 1]))
                            {
                                Console.WriteLine(rowList[i]);
                                tableCell = GetTableCellObject(widthList[i], VerticalAlign.Top,
                                    DocCellMergeEnum.HorizontalStart);
                            }
                            else
                            {
                                tableCell = GetTableCellObject(widthList[i], VerticalAlign.Center);
                            }

                            rowList[i] = rowList[i] ?? "";
                            var text = rowList[i].Split("#$$#");
                            if (text.Length > 1)
                            {
                                for (int j = 0; j < text.Length; j++)
                                {
                                    var run = GetRunObject(false, fontSize);

                                    var paragraph = GetParagraphObject(align[i]);

                                    run.AppendChild(isSerialNumber ? new Text((j + 1) + "）" + text[j]) : new Text(text[j]));

                                    // 将run对象写入段落
                                    paragraph.AppendChild(run);
                                    // 将段落写入单元格
                                    tableCell.AppendChild(paragraph);
                                }
                            }
                            else
                            {
                                foreach (string s in text)
                                {
                                    var run = GetRunObject(false, fontSize);

                                    var paragraph = GetParagraphObject(align[i]);

                                    // 向Run对象中写入数据
                                    run.AppendChild(new Text(s));
                                    // 将run对象写入段落
                                    paragraph.AppendChild(run);
                                    // 将段落写入单元格
                                    tableCell.AppendChild(paragraph);
                                }
                            }


                            //将单元格写入行
                            tableRow.AppendChild(tableCell);
                        }

                        // 将行写入单元格
                        table.AppendChild(tableRow);
                    }
                }
            }
            return table;
        }


        /// <summary>
        /// 获得一个包含垂直合并的表格
        /// 如果为null则进行垂直合并
        /// </summary>
        /// <param name="tableHead">表头</param>
        /// <param name="data">表格数据</param>
        /// <param name="widthList">表格宽度列表，不提供时为100%均分</param>
        /// <param name="fontSize">字体大小，默认为null，为null时继承上一级字体大小</param>
        /// <param name="align">对齐方式，默认左对齐</param>
        /// <param name="isSerialNumber">单元格内换行数据是否编号</param>
        /// <returns></returns>
        public Table GenerateVerticalMergeTable(List<string> tableHead, List<string[]> data, List<int> widthList,
            int? fontSize = null, List<HorizontalAlign> align = null, bool isSerialNumber = false)
        {
            var table = GetTableObject(true, 100);
            if ((tableHead != null) && tableHead.Count != 0)
            {
                TableRow tableRow = new TableRow();
                for (int i = 0; i < tableHead.Count; i++)
                {
                    var tableCell = GetTableCellObject(widthList[i]);

                    var paragraph = GetParagraphObject(HorizontalAlign.Center);

                    var run = GetRunObject(true, fontSize);

                    // 将标题内容添加到run对象中
                    run.AppendChild(new Text(tableHead[i]));
                    // 将run对象添加到段落中
                    paragraph.AppendChild(run);
                    // 将段落对象添加到单元格中
                    tableCell.AppendChild(paragraph);
                    // 将单元格添加到行中
                    tableRow.AppendChild(tableCell);
                }

                // 将表头行添加到表格对象中
                table.AppendChild(tableRow);
            }

            // 如果没有指定对齐方式则初始化为左对齐
            if (align == null)
            {
                align = new List<HorizontalAlign>()
                {
                    HorizontalAlign.Left
                };
            }

            // 补齐对齐方式，如果只有一个则应用于全局，如果指定的对齐不足则使用左对齐补足
            if (align.Count == 1)
            {
                for (int i = 1; i < data[0].Length; i++)
                {
                    align.Add(align[0]);
                }
            }
            else
            {
                for (int i = align.Count; i <= data[0].Length; i++)
                {
                    align.Add(HorizontalAlign.Left);
                }
            }

            // 初始化所有的行
            var tableRowList = new TableRow[data.Count];
            for (int i = 0; i < data.Count; i++)
            {
                tableRowList[i] = new TableRow();
            }
            // 按列循环
            for (int i = 0; i < data[0].Length; i++)
            {
                // 遍历每一行
                for (int j = 0; j < data.Count; j++)
                {
                    TableCell tableCell = null;
                    if (string.IsNullOrEmpty(data[j][i]))
                    {
                        tableCell = GetTableCellObject(widthList[i],
                            cellMerge: DocCellMergeEnum.HorizontalContinue);
                    }
                    // 如果下一单元格没有数据，则开启合并单元格
                    else if ((i + 1) < data[0].Length && string.IsNullOrEmpty(data[j][i + 1]))
                    {
                        tableCell = GetTableCellObject(widthList[i], VerticalAlign.Center,
                            DocCellMergeEnum.HorizontalStart);
                    }
                    else
                    {
                        tableCell = GetTableCellObject(widthList[i], VerticalAlign.Center);
                    }

                    data[j][i] = data[j][i] ?? "";
                    var text = data[j][i].Split("#$$#");
                    if (text.Length > 1)
                    {
                        for (int k = 0; k < text.Length; k++)
                        {
                            var run = GetRunObject(false, fontSize);

                            var paragraph = GetParagraphObject(align[i]);

                            run.AppendChild(isSerialNumber ? new Text((k + 1) + "）" + text[k]) : new Text(text[k]));

                            // 将run对象写入段落
                            paragraph.AppendChild(run);
                            // 将段落写入单元格
                            tableCell.AppendChild(paragraph);
                        }
                    }
                    else
                    {
                        foreach (string s in text)
                        {
                            var run = GetRunObject(false, fontSize);

                            var paragraph = GetParagraphObject(align[i]);

                            // 向Run对象中写入数据
                            run.AppendChild(new Text(s));
                            // 将run对象写入段落
                            paragraph.AppendChild(run);
                            // 将段落写入单元格
                            tableCell.AppendChild(paragraph);
                        }
                    }


                    //将单元格写入行
                    tableRowList[j].AppendChild(tableCell);
                }
            }
            // 将行写入单元格
            foreach (var tableRow in tableRowList)
            {
                table.AppendChild(tableRow);
            }

            return table;
        }


        /// <summary>
        /// 生成表格，通过指定合并起止位置来进行合并
        /// </summary>
        /// <param name="tableHead">表头</param>
        /// <param name="data">表格数据，合并位置也需要提供，可以提供null</param>
        /// <param name="widthList">宽度列表（百分比）</param>
        /// <param name="fontSize">字体大小，不提供或提供为null时继承上文字体大小</param>
        /// <param name="align">对齐方式</param>
        /// <param name="docCellMerges">合并对象列表，用于指定合并的起始位置</param>
        /// <returns></returns>
        public Table GenerateTable(List<string> tableHead, List<string[]> data, List<int> widthList,
            int? fontSize = null, List<HorizontalAlign> align = null, List<DocCellMerge> docCellMerges = null)
        {
            var table = GetTableObject(true, 100);
            var tableCells = new List<List<TableCell>>();

            var horizontalMergesStart = new List<int>();
            var horizontalMergesContinue = new List<int>();
            var verticalMergesStart = new List<int>();
            var verticalMergesContinue = new List<int>();
            if (docCellMerges != null)
            {
                foreach (var docCellMerge in docCellMerges)
                {
                    if (docCellMerge.CellMerge == CellMergeEnum.Horizontal)
                    {
                        horizontalMergesStart.Add(docCellMerge.Start);
                        for (int i = docCellMerge.Start + 1; i <= docCellMerge.Finish; i++)
                        {
                            horizontalMergesContinue.Add(i);
                        }
                    }
                    else
                    {
                        verticalMergesStart.Add(docCellMerge.Start);
                        for (int i = docCellMerge.Start + 1; i <= docCellMerge.Finish; i++)
                        {
                            verticalMergesContinue.Add(i);
                        }
                    }
                }
            }

            // 表头部分
            if ((tableHead != null) && tableHead.Count != 0)
            {
                TableRow tableRow = new TableRow();
                for (int i = 0; i < tableHead.Count; i++)
                {
                    var tableCell = GetTableCellObject(widthList[i]);

                    var paragraph = GetParagraphObject(HorizontalAlign.Center);

                    var run = GetRunObject(true, fontSize);

                    // 将标题内容添加到run对象中
                    run.AppendChild(new Text(tableHead[i]));
                    // 将run对象添加到段落中
                    paragraph.AppendChild(run);
                    // 将段落对象添加到单元格中
                    tableCell.AppendChild(paragraph);
                    // 将单元格添加到行中
                    tableRow.AppendChild(tableCell);
                }

                // 将表头行添加到表格对象中
                table.AppendChild(tableRow);
            }

            // 如果传递的对齐参数与表格列数相同，则用对齐参数列表分别设置每一列对齐方式
            if (align != null && align.Count == data[0].Length)
            {

                for (int i = 0; i < data.Count; i++)
                {
                    TableRow tableRow = new TableRow();
                    // 垂直合并标记
                    var verticalMergeFlag = false;
                    for (int j = 0; j < data[0].Length; j++)
                    {
                        TableCell tableCell = null;
                    }
                }


                foreach (string[] rowList in data)
                {
                    TableRow tableRow = new TableRow();
                    for (int i = 0; i < rowList.Length; i++)
                    {
                        TableCell tableCell = null;
                        // 如果下一单元格没有数据，则开启合并单元格
                        if ((i + 1) < rowList.Length && string.IsNullOrEmpty(rowList[i + 1]))
                        {
                            Console.WriteLine(rowList[i]);
                            tableCell = GetTableCellObject(widthList[i], VerticalAlign.Top,
                                DocCellMergeEnum.HorizontalStart);
                        }
                        else if (string.IsNullOrEmpty(rowList[i]))
                        {
                            tableCell = GetTableCellObject(widthList[i],
                                cellMerge: DocCellMergeEnum.HorizontalContinue);
                        }
                        else
                        {
                            tableCell = GetTableCellObject(widthList[i], VerticalAlign.Center);
                        }

                        rowList[i] = rowList[i] ?? "";
                        var text = rowList[i].Split("#$$#");
                        foreach (string s in text)
                        {
                            var run = GetRunObject(false, fontSize);

                            var paragraph = GetParagraphObject(align[i]);

                            // 向Run对象中写入数据
                            run.AppendChild(new Text(s));
                            // 将run对象写入段落
                            paragraph.AppendChild(run);
                            // 将段落写入单元格
                            tableCell.AppendChild(paragraph);
                        }

                        //将单元格写入行
                        tableRow.AppendChild(tableCell);
                    }

                    // 将行写入单元格
                    table.AppendChild(tableRow);
                }
            }
            else
            {
                throw new Exception("传递的对齐参数与表格列数不相等");
            }

            return table;
        }

        /// <summary>
        /// 将图片插入到Word文档中
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="fileName">图片文件名（全路径）</param>
        /// <param name="imageType">图片类型</param>
        /// <param name="width">图片宽度，单位cm</param>
        /// <param name="height">图片高度，单位cm</param>
        public Run GetPictureRun(WordprocessingDocument doc, string fileName, ImageType imageType, int width = 18, int height = -1)
        {
            if (!File.Exists(fileName))
            {
                throw new Exception($"未在该路径({fileName})找到图片");
            }
            MainDocumentPart mainPart = doc.MainDocumentPart;
            ImagePartType imagePartType = ImagePartType.Png;
            switch (imageType)
            {
                case ImageType.Png:
                    imagePartType = ImagePartType.Png;
                    break;
                case ImageType.Jpeg:
                    imagePartType = ImagePartType.Jpeg;
                    break;
                case ImageType.Bmp:
                    imagePartType = ImagePartType.Bmp;
                    break;
                case ImageType.Gif:
                    imagePartType = ImagePartType.Gif;
                    break;
                case ImageType.Icon:
                    imagePartType = ImagePartType.Icon;
                    break;
            }

            ImagePart imagePart = mainPart.AddImagePart(imagePartType);


            double rate = default;
            // 将图片写入Word
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                imagePart.FeedData(stream);
            }
            // 再次读取计算宽高比
            // 这里因为流写过一次就变成空了，需要深入了解后再优化
            using (FileStream stream = new FileStream(fileName, FileMode.Open))
            {
                Image image = Image.FromStream(stream);
                rate = (double)image.Width / (double)image.Height;
            }


            string relationshipId = mainPart.GetIdOfPart(imagePart);

            long cx = 360000L * width; //360000L = 1厘米
            long cy = default;
            if (height == -1)
            {
                cy = (long)(360000L * width / rate);
            }
            else
            {
                cy = 360000L * height;
            }

            // Define the reference of the image.
            var element =
                 new Drawing(
                     new DW.Inline(
                         new DW.Extent() { Cx = cx, Cy = cy },
                         new DW.EffectExtent()
                         {
                             LeftEdge = 0L,
                             TopEdge = 0L,
                             RightEdge = 0L,
                             BottomEdge = 0L
                         },
                         new DW.DocProperties()
                         {
                             Id = (UInt32Value)1U,
                             Name = Path.GetFileNameWithoutExtension(fileName)
                         },
                         new DW.NonVisualGraphicFrameDrawingProperties(
                             new A.GraphicFrameLocks() { NoChangeAspect = true }),
                         new A.Graphic(
                             new A.GraphicData(
                                 new PIC.Picture(
                                     new PIC.NonVisualPictureProperties(
                                         new PIC.NonVisualDrawingProperties()
                                         {
                                             Id = (UInt32Value)0U,
                                             Name = "New Bitmap Image.jpg"
                                         },
                                         new PIC.NonVisualPictureDrawingProperties()),
                                     new PIC.BlipFill(
                                         new A.Blip(
                                             new A.BlipExtensionList(
                                                 new A.BlipExtension()
                                                 {
                                                     Uri =
                                                        "{28A0092B-C50C-407E-A947-70E740481C1C}"
                                                 })
                                         )
                                         {
                                             Embed = relationshipId,
                                             CompressionState =
                                             A.BlipCompressionValues.Print
                                         },
                                         new A.Stretch(
                                             new A.FillRectangle())),
                                     new PIC.ShapeProperties(
                                         new A.Transform2D(
                                             new A.Offset() { X = 0L, Y = 0L },
                                             new A.Extents() { Cx = cx, Cy = cy }),
                                         new A.PresetGeometry(
                                             new A.AdjustValueList()
                                         )
                                         { Preset = A.ShapeTypeValues.Rectangle }))
                             )
                             { Uri = "http://schemas.openxmlformats.org/drawingml/2006/picture" })
                     )
                     {
                         DistanceFromTop = (UInt32Value)0U,
                         DistanceFromBottom = (UInt32Value)0U,
                         DistanceFromLeft = (UInt32Value)0U,
                         DistanceFromRight = (UInt32Value)0U,
                         EditId = "50D07946"
                     });

            return new Run(element);
        }

    }
}