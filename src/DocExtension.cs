using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KiteDoc
{
    public static class DocExtension
    {
        public static int Replace(this WordprocessingDocument doc,string oldString,string newString)
        {
            // 替换正文中的内容
            var body = doc.MainDocumentPart.Document.Body;
            {
                var paras = body.Descendants<Text>();
                foreach (var text in paras)
                {
                    if (text.Text.Contains(originText))
                    {
                        text.Text = text.Text.Replace(originText, destText);
                    }
                }
            }
            // 替换页脚的内容
            var footer = doc.MainDocumentPart.FooterParts;
            foreach (var footerPart in footer)
            {
                var paras = footerPart.Footer.Descendants<Text>();
                foreach (var text in paras)
                {
                    if (text.Text.Contains(originText))
                    {
                        text.Text = text.Text.Replace(originText, destText);
                    }
                }
            }

            var header = doc.MainDocumentPart.HeaderParts;
            foreach (var headerPart in header)
            {
                var paras = headerPart.Header.Descendants<Text>();
                foreach (var text in paras)
                {
                    if (text.Text.Contains(originText))
                    {
                        text.Text = text.Text.Replace(originText, destText);
                    }
                }
            }
        }
    }
}
