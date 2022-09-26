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
    /// Doc文档Paragraph段落构造器
    /// </summary>
    public class ParagraphBuilder
    {
        private bool buildFlag = false;

        private Paragraph paragraph = new Paragraph();
        private ParagraphProperties paragraphProperties = new ParagraphProperties();
        // 段落样式

        /// <summary>
        /// 设置首行缩进
        /// </summary>
        /// <param name="indentationForTheFirstLine">是否首行缩进</param>
        /// <param name="lineChars">首行缩进值</param>
        /// <returns></returns>
        public ParagraphBuilder SetFirstLineChars(bool indentationForTheFirstLine = true,int lineChars = 200)
        {
            if (paragraphProperties.Indentation == null)
            {
                // 初始化缩进对象
                paragraphProperties.Indentation = new Indentation();
            }

            // 如果需要首行缩进则设置首行缩进
            if (indentationForTheFirstLine)
            {
                // 设置首行缩进
                paragraphProperties.Indentation.FirstLineChars = lineChars;
            }
            else // 不需要则将缩进值置为空
            {
                paragraphProperties.Indentation.FirstLineChars = null;
            }


            return this;
        }

        /// <summary>
        /// 设置段落对齐格式
        /// </summary>
        /// <param name="values">对齐方式</param>
        /// <returns></returns>
        public ParagraphBuilder SetJustification(JustificationValues values)
        {
            paragraphProperties.Justification = new Justification()
            { Val = new EnumValue<JustificationValues>(values) };
            return this;
        }

        /// <summary>
        /// 构造paragraph对象
        /// </summary>
        /// <returns></returns>
        public Paragraph Build()
        {
            if (!buildFlag)
            {
                // 应用样式
                paragraph.AppendChild(paragraphProperties);
                return paragraph;
            }
            else
            {
                throw new Exception("已经构造过对应对象了，如需获得重复对象请使用深克隆");
            }


        }
    }
}
