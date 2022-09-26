﻿using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
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
        /// <summary>
        /// 表格宽度
        /// </summary>
        private TableWidth tableWidth = new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" };

        /// <summary>
        /// 单元格宽度
        /// </summary>
        private TableCellWidth[,] tableCellWidth = new TableCellWidth[0, 0];

        /// <summary>
        /// 表头数据
        /// </summary>
        private string?[] tableHeader = Array.Empty<string>();

        /// <summary>
        /// 内容数据
        /// </summary>
        private string?[,] tableData = new string[0,0];

        /// <summary>
        /// 单元格内段落对齐方式
        /// </summary>
        private JustificationValues[,] justificationValues = new JustificationValues[0, 0];
        /// <summary>
        /// 单元格内垂直对齐方式
        /// </summary>
        private TableVerticalAlignmentValues[,] tableVerticalAlignmentValues = new TableVerticalAlignmentValues[0, 0];

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
        public TableBuilder SetTableCellWidth(TableWidthType tableWidthType, float[] width)
        {
            tableCellWidth = new TableCellWidth[1, width.Length];

            for (int i = 0; i < width.Length; i++)
            {
                tableCellWidth[0,i] = new TableCellWidth(width[i], tableWidthType);
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
            if (width.Length!=0)
            {
                tableCellWidth = new TableCellWidth[width.GetLength(0), width.GetLength(1)];
                for (int i = 0; i < width.GetLength(0); i++)
                {
                    for (int j = 0; j < width.GetLength(1); j++)
                    {
                        tableCellWidth[i,j] = new TableCellWidth(width[i,j],tableWidthType);
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
            justificationValues = new JustificationValues[1,1] { { align } };
            return this;
        }

        /// <summary>
        /// 设置水平对其方式
        /// </summary>
        /// <param name="align">水平对齐方式</param>
        /// <returns></returns>
        public TableBuilder SetJustification(JustificationValues[] align)
        {
            justificationValues = new JustificationValues[1,align.Length];
            // 利用内存复制将一维数组赋值到二维数组
            Buffer.BlockCopy(align, 0, justificationValues, 0, align.Length*sizeof(JustificationValues));
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


        public Table Builder()
        {
            // 添加Table数据

            // 添加Table数据时要将对应的行列的数据样式进行匹配

            // 添加定义的样式（绑定在Table上的表格样式）

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
}