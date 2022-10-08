using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using SixLabors.ImageSharp.Formats.Tiff.Compression.Decompressors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
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
        /// <param name="oldString">被替换的字符串</param>
        /// <param name="table">目标表格</param>
        /// <param name="saveFormatting">是否保留原格式</param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc, string oldString, Table table,bool saveFormatting = true)
        {

            int count = 0;

            // 获得所有的文本元素
            var elements = doc.MainDocumentPart.Document.Descendants<Text>().ToArray();
            // 

            var waitReplace = new List<List<Run>>();
            var continueFlag = false;
            var tempString = new StringBuilder();
            var tempRuns = new List<Run>();
            for (int i = 0; i < elements.Length; i++)
            {

                tempString.Append(elements[i].Text);
                tempRuns.Add((Run)elements[i].Parent);


                // 如果第一次找到开头
                if (!continueFlag && oldString.StartsWith(elements[i].Text))
                {


                    continueFlag = true;

                }
                // 如果添加了之后还是满足开始，说明还可以考虑向后加
                else if (continueFlag && 
                    oldString.StartsWith(tempString.ToString())&&
                    // 只有所有Run的父元素相同的情况下才可以考虑合并文本信息
                    elements[i].Parent.Parent == tempRuns[^1].Parent
                    )
                {
                }
                // 追加了之后不满足了，说明只有前面一部分相同，不是完全相同
                else if (continueFlag)
                {
                    // 将继续标识置为False，下次重新匹配
                    continueFlag = false;

                    tempString.Clear();
                    tempRuns.Clear();
                }
                else
                {
                    tempString.Clear();
                    tempRuns.Clear();
                }


                // 如果匹配成功了，则添加到待替换列表，同时重置待匹配列表和标记
                if (continueFlag&&tempString.ToString().Contains(oldString))
                {
                    continueFlag = false;
                    waitReplace.Add(tempRuns);

                    // 因为后面会添加到数组中，考虑引用传递的问题，这里重新new一个数组
                    tempRuns = new List<Run>();
                    // 清空比对字符串
                    tempString.Clear();
                }
            }

            var tableParentNode = new string[]
            {
                            "body","comment","customXml","docPartBody","endnote","footnote","ftr",
                            "hdr","stdContent","tc"
            };


            // 遍历每一个需要替换为表格文本
            foreach (var item in waitReplace)
            {
                if (saveFormatting)
                {

                    // 得到Text的父对象run
                    var r = item[0];
                    if (r.RunProperties!=null)
                    {
                        var runs = table.Descendants<Run>();
                        foreach (var tblRun in runs)
                        {
                            var rPr = r.RunProperties.Clone() as RunProperties;
                            tblRun.RunProperties = rPr;
                        }
                    }
                }

                // 不能直接添加，要用Clone构造出来
                // 否则会报Cannot insert the OpenXmlElement "newChild" because it is part of a tree.异常
                var copy = table.Clone() as Table;
                if (copy!=null)
                {
                    OpenXmlElement element = item[0].Parent;

                    // Table不能随意插入，父元素只能是以下元素
                    while (!tableParentNode.Contains(element.Parent.XName.LocalName))
                    {
                        element = element.Parent;
                    }


                    //element.Parent.ReplaceChild(copy, element);

                    //foreach (var r in item)
                    //{
                    //    // 找到R的父元素

                    //    r.Remove();
                    //}


                    element.InsertBeforeSelf(copy);

                    element.Remove();

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

            




            // todo: 有时候会因为拼写检查或自定义部分文字样式的时候，会出现文本被截断到多个run中，这时候需要考虑拼接

            //var continueFlag = true;
            //// 使用while是因为替换后会破环foreach，需要再次去获取，否则会替换不完全
            //while (continueFlag)
            //{
            //    // 获取所有的Text对象元素
            //    var elements = doc.MainDocumentPart.Document.Descendants<Text>();
            //    // 默认下次不扫了，如果这次找到了再变更为下次再扫
            //    continueFlag = false;
            //    foreach (var text in elements)
            //    {
            //        

            //        // 找到包含指定文本的text对象
            //        if (text.Text.Contains(oldString))
            //        {
            //            // 如果这次找到了，就再扫描一次
            //            continueFlag = true;

            //            if (saveFormatting)
            //            {

            //                // 得到Text的父对象run
            //                var r = text.Parent as Run;


            //                var runs = table.Descendants<Run>();
            //                foreach (var item in runs)
            //                {
            //                    var rPr = r.RunProperties.Clone() as RunProperties;
            //                    item.RunProperties = rPr;
            //                }
            //            }

            //            // 不能直接添加，要用Clone构造出来
            //            // 否则会报Cannot insert the OpenXmlElement "newChild" because it is part of a tree.异常
            //            var copy = table.Clone() as Table;

            //            OpenXmlElement element = text.Parent;

            //            var tableParentNode = new string[]
            //            {
            //                "body","comment","customXml","docPartBody","endnote","footnote","ftr",
            //                "hdr","stdContent","tc"
            //            };

            //            // Table不能随意插入，父元素只能是以下元素
            //            while (!tableParentNode.Contains( element.Parent.XName.LocalName))
            //            {
            //                element = element.Parent;
            //            }

            //            element.Parent.ReplaceChild(copy, element);

            //            // 如果是在Table中，则至少还要有一个P段落
            //            if (copy.Parent.XName.LocalName == "tc")
            //            {
            //                if (!copy.Parent.Elements<Paragraph>().Any())
            //                {
            //                    copy.InsertAfterSelf(new Paragraph());
            //                }

            //            }

            //            count++;
            //        }
            //    }
            //}


            return count;
        }
    }
}
