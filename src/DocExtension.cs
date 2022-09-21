using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp.Formats.Tiff.Compression.Decompressors;
using System;
using System.Linq;
using System.Xml.Linq;

namespace KiteDoc
{
    /// <summary>
    /// 针对OpenXML提供的一系列扩展方法
    /// </summary>
    public static class DocExtension
    {
        /// <summary>
        /// 替换Word中的字符串
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="oldString">旧文本（被替换的字符串）</param>
        /// <param name="newString">新文本（新的需要的字符串）</param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc,string oldString,string newString)
        {
            // 获取所有的Text对象元素
            var elements = doc.MainDocumentPart.Document.Descendants<Text>();
            int count = 0;
            foreach (var text in elements)
            {
                if (text.Text.Contains(oldString))
                {
                    text.Text = text.Text.Replace(oldString, newString);
                    count++;
                }
            }
            return count;
        }

        /// <summary>
        /// 替换Word中的字符串为表格
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="oldString"></param>
        /// <param name="table"></param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc,string oldString,Table table)
        {
            
            return Replace(doc,oldString,table,true);
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="doc"></param>
        /// <param name="oldString"></param>
        /// <param name="table"></param>
        /// <param name="saveFormatting"></param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc, string oldString, Table table,bool saveFormatting)
        {

            int count = 0;
            var continueFlag = true;
            // 使用while是因为替换后会破环foreach，需要再次去获取，否则会替换不完全
            while (continueFlag)
            {
                // 获取所有的Text对象元素
                var elements = doc.MainDocumentPart.Document.Descendants<Text>();
                // 默认下次不扫了，如果这次找到了再变更为下次再扫
                continueFlag = false;
                foreach (var text in elements)
                {
                    // 找到包含指定文本的text对象
                    if (text.Text.Contains(oldString))
                    {
                        // 如果这次找到了，就再扫描一次
                        continueFlag = true;

                        if (saveFormatting)
                        {

                            // 得到Text的父对象run
                            var r = text.Parent as Run;


                            var runs = table.Descendants<Run>();
                            foreach (var item in runs)
                            {
                                var rPr = r.RunProperties.Clone() as RunProperties;
                                item.RunProperties = rPr;
                            }
                        }

                        // 不能直接添加，要用Clone构造出来
                        // 否则会报Cannot insert the OpenXmlElement "newChild" because it is part of a tree.异常
                        var copy = table.Clone() as Table;

                        OpenXmlElement element = text.Parent;

                        var tableParentNode = new string[]
                        {
                            "body","comment","customXml","docPartBody","endnote","footnote","ftr",
                            "hdr","stdContent","tc"
                        };

                        // Table不能随意插入，父元素只能是以下元素
                        while (!tableParentNode.Contains( element.Parent.XName.LocalName))
                        {
                            element = element.Parent;
                        }

                        element.Parent.ReplaceChild(copy, element);

                        // 如果是在Table中，则至少还要有一个P段落
                        if (copy.Parent.XName.LocalName == "tc")
                        {
                            if (!copy.Parent.Elements<Paragraph>().Any())
                            {
                                copy.InsertAfterSelf(new Paragraph());
                            }
                            
                        }

                        count++;
                    }
                }
            }
            
            
            return count;
        }
    }
}
