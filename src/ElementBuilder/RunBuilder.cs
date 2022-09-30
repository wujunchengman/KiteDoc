using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KiteDoc.ElementBuilder
{
    /// <summary>
    /// 构建一个Doc的Run对象
    /// </summary>
    public class RunBuilder
    {
        /// <summary>
        /// Run对象
        /// </summary>
        private Run run = new();
        /// <summary>
        /// Run属性信息
        /// </summary>
        private RunProperties runProperties = new();

        /// <summary>
        /// 设置字体加粗
        /// </summary>
        /// <param name="isBold">是否加粗</param>
        /// <returns></returns>
        public RunBuilder SetBold(bool isBold = true)
        {
            if (isBold)
            {
                runProperties.Bold = new Bold();
            }
            else
            {
                runProperties.Bold = null;
            }
            return this;
        }

        /// <summary>
        /// 设置字体大小
        /// </summary>
        /// <param name="fontSize">字体大小</param>
        /// <returns></returns>
        public RunBuilder SetFontSize(float? fontSize)
        {
            if (fontSize != null)
            {
                runProperties.FontSize =new FontSize() { Val = ((int)(fontSize * 2+0.5)).ToString() };
            }

            return this;
        }


        /// <summary>
        /// 设置文字字体
        /// </summary>
        /// <param name="zhFont">中文字体名</param>
        /// <param name="enFont">英文字体名</param>
        /// <returns></returns>
        public RunBuilder SetFont(string zhFont,string? enFont = null)
        {
            if (string.IsNullOrWhiteSpace(zhFont))
            {
                if (enFont == null)
                {
                    enFont = zhFont;
                }

                runProperties.RunFonts = new RunFonts()
                {
                    Ascii = enFont,
                    HighAnsi = enFont,
                    // 复杂文种
                    ComplexScript = zhFont,
                    // 中文
                    EastAsia = zhFont,
                };
            }
            return this;
        }

        /// <summary>
        /// 向Run中添加文本
        /// </summary>
        /// <param name="text">文本内容</param>
        /// <returns></returns>
        public RunBuilder AppendText(string text)
        {
            run.AppendChild(new Text(text));
            return this;
        }

        /// <summary>
        /// 构建Run对象
        /// </summary>
        /// <returns></returns>
        public Run Build()
        {
            run.RunProperties =runProperties;

            return run;
        }
    }
}
