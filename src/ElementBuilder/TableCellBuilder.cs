using DocumentFormat.OpenXml.Wordprocessing;
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

        public TableCell Build()
        {
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
}
