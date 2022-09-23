using DocumentFormat.OpenXml;
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
        /// 设置表格框线
        /// </summary>
        /// <param name="tableborderScope">边框范围</param>
        /// <param name="borderSize">边框大小</param>
        /// <param name="borderType">框线类型</param>
        /// <returns></returns>
        public TableBuilder SetBorder(TableBorderScope tableborderScope = TableBorderScope.All,float borderSize = 1,BorderValues borderType = BorderValues.Single)
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
                        tableBorders.AppendChild(new TopBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal });
                        tableBorders.AppendChild(new BottomBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal });
                        tableBorders.AppendChild(new LeftBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal });
                        tableBorders.AppendChild(new RightBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal });
                        tableBorders.AppendChild(new InsideHorizontalBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal });
                        tableBorders.AppendChild(new InsideVerticalBorder { Val = new EnumValue<BorderValues>(borderType), Size = borderSizeVal });

                    }
                    break;
                case TableBorderScope.Left:
                    {
                        var left = tableBorders.Elements<LeftBorder>().ToList();
                        if (left.Any())
                        {
                            
                        }

                    }
                    break;
                case TableBorderScope.Right:
                    break;
                case TableBorderScope.Top:
                    break;
                case TableBorderScope.Bottom:
                    break;
                default:
                    break;
            }
            return this;
        }
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
        Bottom
    }
}
