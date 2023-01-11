using DocumentFormat.OpenXml;
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
    /// 表格（具体的格子）构造器
    /// </summary>
    public class TableCellBuilder
    {
        private TableCell tableCell = new();
        private TableCellProperties tableCellProperties = new TableCellProperties();
        private TableCellWidth? tableCellWidth;
        private TableVerticalAlignmentValues tableVerticalAlignmentValues = TableVerticalAlignmentValues.Center;
        private TableCellMerge tableCellMerge = TableCellMerge.Normal;
        /// <summary>
        /// 设置单元格宽度
        /// </summary>
        /// <param name="tableCellWidth">单元格宽度</param>
        /// <returns></returns>
        public TableCellBuilder SetTableCellWidth(TableCellWidth tableCellWidth)
        {
            this.tableCellWidth = tableCellWidth;
            return this;
        }

        /// <summary>
        /// 设置表格合并标记
        /// </summary>
        /// <param name="tableCellMerge">表格合并标记</param>
        /// <returns></returns>
        public TableCellBuilder SetTableCellMerge(TableCellMerge tableCellMerge)
        {
            this.tableCellMerge = tableCellMerge;
            return this;
        }

        /// <summary>
        /// 构造表格单元格
        /// </summary>
        /// <returns></returns>
        public TableCell Build()
        {
            // 设置表格合并
            // 添加合并单元格标记
            switch (tableCellMerge)
            {
                case TableCellMerge.Normal:
                    break;
                case TableCellMerge.HorizontalStart:
                    tableCellProperties.AppendChild(new HorizontalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    break;
                case TableCellMerge.HorizontalContinue:
                    //Console.WriteLine("continue");
                    tableCellProperties.AppendChild(new HorizontalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Continue) });
                    break;
                case TableCellMerge.VerticalStart:
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    break;
                case TableCellMerge.VerticalContinue:
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Continue) });
                    break;
                case TableCellMerge.HorizontalAndVerticalStart:
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    tableCellProperties.AppendChild(new VerticalMerge()
                    { Val = new EnumValue<MergedCellValues>(MergedCellValues.Restart) });
                    break;
            }



            // 设置宽度
            if (tableCellWidth!=null)
            {
                var width =  new DocumentFormat.OpenXml.Wordprocessing.TableCellWidth();
                switch (tableCellWidth.TableWidthType)
                {
                    case TableWidthType.Percent:
                        width.Type = TableWidthUnitValues.Pct;
                        width.Width = (Math.Round(tableCellWidth.Width, 1) * 50).ToString();
                        break;
                    case TableWidthType.Cm:
                        // Dxa是二十分之一点，其与厘米的换算需要先将cm转换为像素，然后乘以20，Word使用72dpi显示

                        // Dxa =  cm/2.54*72*20

                        // Word中的行为是先计算到像素点，然后保留一位小数，再乘以20，得到二十分之一点值
                        width.Type = TableWidthUnitValues.Dxa;
                        width.Width = (Math.Round((tableCellWidth.Width * 72 / 2.54), 1) * 20).ToString();
                        break;
                    case TableWidthType.Nil:
                        width.Type = TableWidthUnitValues.Nil;
                        break;
                    case TableWidthType.Auto:
                        width.Type = TableWidthUnitValues.Auto;
                        break;
                    default:
                        break;
                }
                tableCellProperties.TableCellWidth = width;
            }

            // 设置居中
            tableCellProperties.TableCellVerticalAlignment = new TableCellVerticalAlignment { Val = tableVerticalAlignmentValues };
            

            // 绑定属性
            tableCell.TableCellProperties = tableCellProperties;

            // 返回表格
            return tableCell;
        }

    }

    /// <summary>
    /// 表格合并方式
    /// </summary>
    public enum TableCellMerge
    {
        /// <summary>
        /// 普通
        /// </summary>
        Normal,
        /// <summary>
        /// 水平合并起始
        /// </summary>
        HorizontalStart,
        /// <summary>
        /// 水平合并单元格(该单元格不会被显示)
        /// </summary>
        HorizontalContinue,
        /// <summary>
        /// 垂直合并单元格起始
        /// </summary>
        VerticalStart,
        /// <summary>
        /// 垂直合并单元格(该单元格不会被显示)
        /// </summary>
        VerticalContinue,
        /// <summary>
        /// 水平合并和垂直合并共同的起始
        /// </summary>
        HorizontalAndVerticalStart,
        /// <summary>
        /// 水平合并和垂直合并共同的合并单元格(该单元格不会被显示)
        /// </summary>
        HorizontalAndVerticalContinue
    }
}
