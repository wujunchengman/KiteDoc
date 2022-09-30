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
    /// Doc文档Paragraph段落构造器
    /// </summary>
    public class ParagraphBuilder
    {
        /// <summary>
        /// 是否已经构建过了
        /// </summary>
        private bool buildFlag = false;

        /// <summary>
        /// 段落对象
        /// </summary>
        private Paragraph paragraph = new Paragraph();

        /// <summary>
        /// 段落样式
        /// </summary>
        private ParagraphProperties paragraphProperties = new ParagraphProperties();

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
        /// 添加普通文本到段落
        /// </summary>
        /// <param name="text">内容文本</param>
        /// <param name="isBold">是否加粗</param>
        /// <param name="fontSize">文字大小</param>
        /// <param name="font">字体</param>
        /// <returns></returns>
        public ParagraphBuilder AppendText(string text,bool isBold = false,float? fontSize = null,string font = "宋体")
        {
            var run = new RunBuilder()
                .SetFont(font)
                .SetFontSize(fontSize)
                .SetBold(isBold)
                .AppendText(text)
                .Build();
            paragraph.AddChild(run);
            return this;
        }

        /// <summary>
        /// 添加Run对象到段落中
        /// </summary>
        /// <param name="run">Doc文档的Run对象</param>
        /// <returns></returns>
        public ParagraphBuilder AppendRun(Run run)
        {
            paragraph.AddChild(run);
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
                paragraph.ParagraphProperties = paragraphProperties;
                return paragraph;
            }
            else
            {
                throw new Exception("已经构造过对应对象了，如需获得重复对象请使用深克隆");
            }


        }
    }
}
