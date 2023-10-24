using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.Enum;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KiteDoc.ElementBuilder
{
    /// <summary>
    /// Doc文档Table组件构造器
    /// </summary>
    public class TableBuilder
    {
        private Table table = new Table();
        private TableProperties tableProperties = new TableProperties();
        private TableBorders tableBorders = new TableBorders();
        private TableCellMergeType tableCellMerge = TableCellMergeType.None;

        /// <summary>
        /// 表格宽度
        /// </summary>
        private TableWidth tableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" };

        /// <summary>
        /// 单元格宽度
        /// </summary>
        private TableCellWidth[,] tableCellWidth = new TableCellWidth[0, 0];

        /// <summary>
        /// 表头单元格宽度
        /// </summary>
        private TableCellWidth[] tableHeaderCellWidth = Array.Empty<TableCellWidth>();

        /// <summary>
        /// 表头数据
        /// </summary>
        private List<string?> tableHeader = new();

        /// <summary>
        /// 表头水平对齐样式方式
        /// </summary>
        private JustificationValues tableHeaderJustificationValue = JustificationValues.Center;

        /// <summary>
        /// 内容数据
        /// </summary>
        private List<List<string?>> tableData = new();


        /// <summary>
        /// 表格字体大小
        /// </summary>
        private float? tableDataFontSize = null;

        /// <summary>
        /// 单元格内段落对齐方式
        /// </summary>
        private JustificationValues[,] justificationValues = new JustificationValues[0, 0];
        /// <summary>
        /// 单元格内垂直对齐方式
        /// </summary>
        private TableVerticalAlignmentValues[,] tableVerticalAlignmentValues = new TableVerticalAlignmentValues[0, 0];

        /// <summary>
        /// 分割段落字符串
        /// </summary>
        private string? splitString;

        /// <summary>
        /// 是否对分割段落进行编号
        /// </summary>
        private bool isSerialNumber;

        /// <summary>
        /// 设置表格框线
        /// </summary>
        /// <param name="tableborderScope">边框范围</param>
        /// <param name="borderSize">边框大小</param>
        /// <param name="borderType">框线类型</param>
        /// <returns></returns>
        public TableBuilder SetBorder(TableBorderScope tableborderScope = TableBorderScope.All, float borderSize = 1, BorderValues borderType = BorderValues.Single)
        {
            // 将边框线换算为Excel中的值，Excel中2为1榜
            var borderSizeVal = (uint)(borderSize * 2 + 0.5);

            switch (tableborderScope)
            {
                case TableBorderScope.None:
                    {
                        // 移除边框线
                        tableBorders.RemoveAllChildren();
                    }
                    break;
                case TableBorderScope.All:
                    {
                        // 先移除已有的框线设置
                        tableBorders.RemoveAllChildren();

                        // 将所有框线添加到到表格框线中
                        tableBorders.TopBorder = new TopBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                        tableBorders.BottomBorder = new BottomBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                        tableBorders.LeftBorder = new LeftBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                        tableBorders.RightBorder = new RightBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                        tableBorders.InsideHorizontalBorder = new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                        tableBorders.InsideVerticalBorder = new InsideVerticalBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };

                    }
                    break;
                case TableBorderScope.Left:
                    {
                        tableBorders.LeftBorder = new LeftBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                    }
                    break;
                case TableBorderScope.Right:
                    {
                        tableBorders.RightBorder = new RightBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                    }
                    break;
                case TableBorderScope.Top:
                    {
                        tableBorders.TopBorder = new TopBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                    }
                    break;
                case TableBorderScope.Bottom:
                    {
                        tableBorders.BottomBorder = new BottomBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                    }
                    break;
                case TableBorderScope.InsideHorizontal:
                    {
                        tableBorders.InsideHorizontalBorder = new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                    }
                    break;
                case TableBorderScope.InsideVertical:
                    {
                        tableBorders.InsideVerticalBorder = new InsideVerticalBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal };
                    }
                    break;
                default:
                    break;
            }
            return this;
        }

        /// <summary>
        /// 配置表格宽度
        /// </summary>
        /// <param name="tableWidthType">表格宽度类型</param>
        /// <param name="width">宽度值</param>
        /// <returns></returns>
        public TableBuilder SetTableWidth(TableWidthType tableWidthType, float width)
        {
            switch (tableWidthType)
            {
                case TableWidthType.Percent:
                    // Word中设置百分比的宽度行为是先将百分比保留一位小数，然后乘以五十得到五十分之一百分比值
                    tableWidth = new TableWidth { Type = TableWidthUnitValues.Pct, Width = (Math.Round(width, 1) * 50).ToString() };
                    break;
                case TableWidthType.Cm:
                    // Dxa是二十分之一点，其与厘米的换算需要先将cm转换为像素，然后乘以20，Word使用72dpi显示

                    // Dxa =  cm/2.54*72*20

                    // Word中的行为是先计算到像素点，然后保留一位小数，再乘以20，得到二十分之一点值
                    tableWidth = new TableWidth { Type = TableWidthUnitValues.Dxa, Width = (Math.Round((width * 72 / 2.54), 1) * 20).ToString() };
                    break;
                case TableWidthType.Nil:
                    tableWidth = new TableWidth { Type = TableWidthUnitValues.Nil };
                    break;
                case TableWidthType.Auto:
                    tableWidth = new TableWidth { Type = TableWidthUnitValues.Auto };
                    break;
                default:
                    break;
            }
            return this;
        }

        /// <summary>
        /// 设置表格单元格宽度
        /// </summary>
        /// <param name="tableWidthType">宽度计算类型</param>
        /// <param name="width">宽度值</param>
        /// <returns></returns>
        public TableBuilder SetTableCellWidth(TableWidthType tableWidthType, float width)
        {
            tableCellWidth = new TableCellWidth[1, 1] { { new TableCellWidth(width, tableWidthType) } };
            return this;

        }


        /// <summary>
        /// 设置表格单元格宽度
        /// </summary>
        /// <param name="tableWidthType"></param>
        /// <param name="width"></param>
        /// <returns></returns>
        public TableBuilder SetTableCellWidth(TableWidthType tableWidthType,float[] width)
        {
            tableCellWidth = new TableCellWidth[1, width.Length];

            for (int i = 0; i < width.Length; i++)
            {
                tableCellWidth[0, i] = new TableCellWidth(width[i], tableWidthType);
            }

            return this;
        }

        /// <summary>
        /// 设置表格单元格宽度
        /// </summary>
        /// <param name="tableWidthType">宽度计算类型</param>
        /// <param name="width">宽度值</param>
        /// <returns></returns>
        public TableBuilder SetTableCellWidth(TableWidthType tableWidthType, float[,] width)
        {
            if (width.Length != 0)
            {
                tableCellWidth = new TableCellWidth[width.GetLength(0), width.GetLength(1)];
                for (int i = 0; i < width.GetLength(0); i++)
                {
                    for (int j = 0; j < width.GetLength(1); j++)
                    {
                        tableCellWidth[i, j] = new TableCellWidth(width[i, j], tableWidthType);
                    }
                }

            }
            return this;
        }

        /// <summary>
        /// 设置水平对其方式
        /// </summary>
        /// <param name="align">水平对齐方式</param>
        /// <returns></returns>
        public TableBuilder SetJustification(JustificationValues align = JustificationValues.Center)
        {
            justificationValues = new JustificationValues[1, 1] { { align } };
            return this;
        }

        /// <summary>
        /// 设置水平对其方式
        /// </summary>
        /// <param name="align">水平对齐方式</param>
        /// <returns></returns>
        public TableBuilder SetJustification(JustificationValues[] align)
        {
            justificationValues = new JustificationValues[1, align.Length];
            // 利用内存复制将一维数组赋值到二维数组
            Buffer.BlockCopy(align, 0, justificationValues, 0, align.Length * sizeof(JustificationValues));
            return this;
        }

        /// <summary>
        /// 设置水平对其方式
        /// </summary>
        /// <param name="align">水平对齐方式</param>
        /// <returns></returns>
        public TableBuilder SetJustification(JustificationValues[,] align)
        {
            justificationValues = align;
            return this;
        }

        /// <summary>
        /// 设置表格标题
        /// </summary>
        /// <param name="header">标题文字</param>
        /// <returns></returns>
        public TableBuilder SetTableHeader(List<string?> header)
        {
            tableHeader = header;
            return this;
        }

        /// <summary>
        /// 设置表头水平对齐方式
        /// </summary>
        /// <param name="align">水平对齐方式</param>
        /// <returns></returns>
        public TableBuilder SetTableHeaderJustification(JustificationValues align = JustificationValues.Center)
        {
            tableHeaderJustificationValue = align;
            return this;
        }

        /// <summary>
        /// 设置表格标题单元格宽度
        /// </summary>
        /// <param name="tableWidthType">宽度类型</param>
        /// <param name="width">宽度值</param>
        /// <returns></returns>
        public TableBuilder SetTableHeaderCellWidth(TableWidthType tableWidthType, float[] width)
        {
            tableHeaderCellWidth = new TableCellWidth[width.Length];

            for (int i = 0; i < width.Length; i++)
            {
                tableHeaderCellWidth[i] = new TableCellWidth(width[i], tableWidthType);
            }
            return this;
        }

        /// <summary>
        /// 设置表格文字
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        public TableBuilder SetTableData(List<List<string?>> data)
        {
            tableData = data;
            return this;
        }

        public TableBuilder SetTableDataFontSize(float fontSize)
        {
            tableDataFontSize = fontSize;
            return this;
        }

        /// <summary>
        /// 设置表格内文字段落换行
        /// </summary>
        /// <param name="splitString">分割字符换</param>
        /// <param name="isSerialNumber">是否对分割后的段落编号</param>
        /// <returns></returns>
        public TableBuilder SetDataSplitParagraph(string splitString, bool isSerialNumber)
        {
            this.splitString = splitString;
            this.isSerialNumber = isSerialNumber;
            return this;
        }


        // todo: 设置水平合并
        /// <summary>
        /// 设置水平方向的Null内容合并
        /// </summary>
        /// <param name="merge"></param>
        /// <returns></returns>
        public TableBuilder SetHorizationNullMerge(bool merge = true)
        {
            if (merge)
            {
                tableCellMerge = TableCellMergeType.HorizontalNullMerge;
            }
            return this;
        }

        // todo: 设置垂直合并

        /// <summary>
        /// 生成Table对象
        /// </summary>
        /// <returns></returns>
        public Table Build()
        {


            // 添加Table数据时要将对应的行列的数据样式进行匹配

            // 添加定义的样式（绑定在Table上的表格样式）
            tableProperties.TableBorders = tableBorders;
            tableProperties.TableWidth = tableWidth;

            var tblPr = table.Elements<TableProperties>();
            foreach (var item in tblPr)
            {
                item.Remove();
            }

            // 样式属性必须在子元素的第一个
            table.AppendChild(tableProperties);

            // 如果没有指定表头宽度则全部指定为自动
            if (tableHeaderCellWidth.Length == 0)
            {
                tableHeaderCellWidth = new TableCellWidth[tableHeader.Count];
                for (int i = 0; i < tableHeader.Count; i++)
                {
                    tableHeaderCellWidth[i] = new TableCellWidth(0, TableWidthType.Auto);
                }
            }

            // 添加Table表头数据
            if ((tableHeader != null) && tableHeader.Count != 0)
            {
                TableRow tableRow = new TableRow();
                for (int i = 0; i < tableHeader.Count; i++)
                {
                    // 获得一个单元格，需要判断单元格是否需要合并，怎么合并
                    var tableCell = new TableCellBuilder().SetTableCellWidth(tableHeaderCellWidth[i]).Build();
                    //var tableCell = GetTableCellObject(widthList[i]);

                    var paragraph = new ParagraphBuilder()
                        .AppendText(tableHeader[i])
                        .SetJustification(tableHeaderJustificationValue)
                        .Build();


                    // 将段落对象添加到单元格中
                    tableCell.AppendChild(paragraph);
                    // 将单元格添加到行中
                    tableRow.AppendChild(tableCell);
                }

                // 将表头行添加到表格对象中
                table.AppendChild(tableRow);
            }


            if (tableData.Count > 0)
            {
                var rowsCount = tableData.Count;
                var colCount = tableData[0].Count;

                // 将表格宽度配置统一成统一格式
                {
                    // 可能没有指定宽度
                    if (tableCellWidth.Length == 0)
                    {
                        // 重设表格宽度大小
                        tableCellWidth = new TableCellWidth[rowsCount, colCount];

                        // 如果设置了表头宽度
                        if (tableHeaderCellWidth.Length != 0)
                        {

                            for (int i = 0; i < rowsCount; i++)
                            {
                                for (int j = 0; j < colCount; j++)
                                {
                                    tableCellWidth[i, j] = tableHeaderCellWidth[j];
                                }
                            }
                        }
                        else
                        {
                            for (int i = 0; i < rowsCount; i++)
                            {
                                for (int j = 0; j < colCount; j++)
                                {
                                    tableCellWidth[i, j] = new TableCellWidth(0, TableWidthType.Auto);
                                }
                            }
                        }
                    }
                    // 可能指定了一行宽度
                    else if (tableCellWidth.GetLength(0) == 1 && tableCellWidth.GetLength(1) == colCount)
                    {
                        // 重设表格宽度大小
                        var newTableCellWidth = new TableCellWidth[rowsCount, colCount];
                        for (int i = 0; i < rowsCount; i++)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                newTableCellWidth[i, j] = tableCellWidth[0, j];
                            }
                        }

                        tableCellWidth = newTableCellWidth;

                    }
                    // 可能指定了单个宽度
                    else
                    {
                        // todo: 这里有问题，现在是指定了表格宽度就会抛异常
                        throw new ArgumentException("暂未支持的表格宽度设置方案");
                    }
                    

                    // 可能指定了详细宽度
                }

                // 将表格对齐方式配置格式统一
                {
                    // 没有设置的情况
                    if (justificationValues.Length == 0)
                    {
                        justificationValues = new JustificationValues[rowsCount, colCount];
                        for (int i = 0; i < rowsCount; i++)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                justificationValues[i, j] = JustificationValues.Left;
                            }
                        }
                    }
                    // 设置了所有单元格统一对齐方式的情况
                    else if (justificationValues.Length == 1)
                    {
                        var val = justificationValues[0, 0];
                        justificationValues = new JustificationValues[rowsCount, colCount];
                        for (int i = 0; i < rowsCount; i++)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                justificationValues[i, j] = val;
                            }
                        }
                    }
                    // 只有一行，一行有多个值的情况
                    else if (justificationValues.GetLength(0) == 1 && justificationValues.GetLength(1) > 1)
                    {
                        var val = new JustificationValues[rowsCount, colCount];
                        for (int i = 0; i < rowsCount; i++)
                        {
                            for (int j = 0; j < colCount; j++)
                            {
                                val[i, j] = justificationValues[0, j];
                            }
                        }
                        justificationValues = val;
                    }
                    else
                    {
                        // 已经是详细配置的不再处理
                    }
                }

                // 如果制定了分割字符串就对内容进行分割判断
                if (false)
                {
                    for (int i = 0; i < tableData.Count; i++)
                    {
                        TableRow tableRow = new TableRow();
                        for (int j = 0; j < tableData[0].Count; j++)
                        {
                            // 获得一个单元格，需要判断单元格是否需要合并，怎么合并
                            var tableCell = new TableCellBuilder()
                                .SetTableCellWidth(tableCellWidth[i, j])
                                .Build();

                            var split = tableData[i][j]?.Split(splitString);

                            if (split != null)
                            {
                                for (int k = 0; k < split.Length; k++)
                                {
                                    Paragraph paragraph;
                                    if (isSerialNumber)
                                    {
                                        paragraph = new ParagraphBuilder()
                                           .AppendText((k + 1) + ". " + split[k],fontSize:tableDataFontSize)
                                           .SetJustification(justificationValues[i, j])
                                           .Build();
                                    }
                                    else
                                    {
                                        paragraph = new ParagraphBuilder()
                                           .AppendText(split[k],fontSize: tableDataFontSize)
                                           .SetJustification(justificationValues[i, j])
                                           .Build();
                                    }



                                    // 将段落对象添加到单元格中
                                    tableCell.AppendChild(paragraph);
                                }
                            }



                            // 将单元格添加到行中
                            tableRow.AppendChild(tableCell);
                        }

                        table.AppendChild(tableRow);
                    }
                }
                else
                {
                    for (int i = 0; i < tableData.Count; i++)
                    {
                        TableRow tableRow = new TableRow();
                        for (int j = 0; j < tableData[0].Count; j++)
                        {
                            var tableCellBuilder = new TableCellBuilder();

                            // 如果是水平Null值合并
                            if (tableCellMerge == TableCellMergeType.HorizontalNullMerge)
                            {
                                // 如果是null值，则是继续合并
                                if (tableData[i][j]==null)
                                {
                                    tableCellBuilder.SetTableCellMerge(TableCellMerge.HorizontalContinue);
                                }
                                // 如果下一单元格没有数据，则开启合并单元格
                                else if ((j + 1) < tableData[i].Count && tableData[i][j + 1]==null)
                                {
                                    tableCellBuilder.SetTableCellMerge(TableCellMerge.HorizontalStart);
                                }
                                else
                                {
                                    // 普通情况不需要合并
                                }
                            }

                            var tableCell = tableCellBuilder
                                .SetTableCellWidth(tableCellWidth[i, j])
                                .Build();

                            if (splitString!=null)
                            {
                                var split = tableData[i][j]?.Split(splitString);

                                if (split != null)
                                {
                                    for (int k = 0; k < split.Length; k++)
                                    {
                                        Paragraph paragraph;
                                        if (isSerialNumber)
                                        {
                                            paragraph = new ParagraphBuilder()
                                                .AppendText((k + 1) + ". " + split[k],fontSize:tableDataFontSize)
                                                .SetJustification(justificationValues[i, j])
                                                .Build();
                                        }
                                        else
                                        {
                                            paragraph = new ParagraphBuilder()
                                                .AppendText(split[k],fontSize: tableDataFontSize)
                                                .SetJustification(justificationValues[i, j])
                                                .Build();
                                        }
                                        // 将段落对象添加到单元格中
                                        tableCell.AppendChild(paragraph);
                                    }
                                }
                                else
                                {
                                    // 当当前位置是null时，TableCell中也要放一个空的Paragraph，以符合文档规范
                                    
                                    var paragraph = new ParagraphBuilder()
                                        .AppendText(tableData[i][j], fontSize: tableDataFontSize)
                                        .SetJustification(justificationValues[i, j])
                                        .Build();
                                    
                                    // 将段落对象添加到单元格中
                                    tableCell.AppendChild(paragraph);
                                }
                                

                            }
                            else
                            {
                                var paragraph = new ParagraphBuilder()
                                    .AppendText(tableData[i][j], fontSize: tableDataFontSize)
                                    .SetJustification(justificationValues[i, j])
                                    .Build();


                                // 将段落对象添加到单元格中
                                tableCell.AppendChild(paragraph);
                            }
                            
                            // 将单元格添加到行中
                            tableRow.AppendChild(tableCell);
                        }

                        table.AppendChild(tableRow);
                    }
                }



            }


            return table;


        }

    }

    /// <summary>
    /// 单元格宽度
    /// </summary>
    public class TableCellWidth
    {
        /// <summary>
        /// 初始化单元格宽度对象
        /// </summary>
        /// <param name="width">宽度</param>
        /// <param name="tableWidthType">宽度计算类型</param>
        public TableCellWidth(float width, TableWidthType tableWidthType)
        {
            Width = width;
            TableWidthType = tableWidthType;
        }


        /// <summary>
        /// 表格单元格宽度
        /// </summary>
        public float Width { get; set; }
        /// <summary>
        /// 表格单元格宽度计算方式
        /// </summary>
        public TableWidthType TableWidthType { get; set; }
    }

    /// <summary>
    /// 水平对齐方式
    /// </summary>
    public enum HorizontalAlign
    {
        /// <summary>
        /// 左对齐
        /// </summary>
        Left,
        /// <summary>
        /// 居中对齐
        /// </summary>
        Center,
        /// <summary>
        /// 右对齐
        /// </summary>
        Right
    }


    /// <summary>
    /// 表格宽度类型
    /// </summary>
    public enum TableWidthType
    {
        /// <summary>
        /// 百分比
        /// </summary>
        Percent,
        /// <summary>
        /// 厘米
        /// </summary>
        Cm,
        /// <summary>
        /// 无宽度
        /// </summary>
        Nil,
        /// <summary>
        /// 自动决定宽度
        /// </summary>
        Auto,

    }

    /// <summary>
    /// 表格边框范围
    /// </summary>
    public enum TableBorderScope
    {
        /// <summary>
        /// 无边框
        /// </summary>
        None,
        /// <summary>
        /// 所有框线
        /// </summary>
        All,
        /// <summary>
        /// 左框线
        /// </summary>
        Left,
        /// <summary>
        /// 右框线
        /// </summary>
        Right,
        /// <summary>
        /// 上框线
        /// </summary>
        Top,
        /// <summary>
        /// 下框线
        /// </summary>
        Bottom,
        /// <summary>
        /// 内部水平框线
        /// </summary>
        InsideHorizontal,
        /// <summary>
        /// 内部垂直框线
        /// </summary>
        InsideVertical

    }

    /// <summary>
    /// 表格合并方式
    /// </summary>
    public enum TableCellMergeType
    {
        /// <summary>
        /// 不合并
        /// </summary>
        None,
        /// <summary>
        /// 水平Null合并
        /// </summary>
        HorizontalNullMerge,

    }

}
