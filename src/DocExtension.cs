using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.Utils;
using System;
using System.Collections.Generic;
using System.Linq;

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
            var elements = doc.FindAllTextElement();

            return ReplaceString(elements,oldString,newString);
        }

        /// <summary>
        /// 批量替换Word中的字符串
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="keyValuePairs">替换键值对，Key为旧字符串，value为新字符串</param>
        /// <returns></returns>
        public static List<int> Replace(this WordprocessingDocument doc, IEnumerable<KeyValuePair<string,string>> keyValuePairs)
        {
            var result = new List<int>(keyValuePairs.Count());
            // 获取所有的Text对象元素
            var elements = doc.FindAllTextElement();
            foreach (var item in keyValuePairs)
            {
                result.Add( ReplaceString(elements, item.Key, item.Value));
            }
            return result;
        }

        private static int ReplaceString(Text[]? elements, string oldString,string newString)
        {
            int count = 0;
            var waitReplace = DocElementUtils.FindRun(elements, oldString);
            foreach (var item in waitReplace)
            {
                if (item.Count == 1)
                {
                    var text = item[0].Descendants<Text>().FirstOrDefault();
                    text.Text = text.Text.Replace(oldString, newString);
                }
                else
                {
                    var text = item[0].Descendants<Text>().FirstOrDefault();

                    // 把所有Run的Text拼到一起，然后去替换对应的文本
                    var dest = string.Concat(item.Select(x => x.Descendants<Text>().FirstOrDefault()).Select(t => t.Text));

                    text.Text = dest.Replace(oldString, newString);

                    // todo: 需要考虑最后一个Run对象中是否还有其他的文本
                    foreach (var run in item.Skip(1))
                    {
                        run.Remove();
                    }
                }
                count++;

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

            // 获得所有的文本元素
            var elements = doc.FindAllTextElement();

            return ReplaceTable(elements, oldString, table, saveFormatting);

            
        }

        /// <summary>
        /// 批量替换Word中的字符串为表格
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="replaces">替换参数组</param>
        /// <returns></returns>
        public static List<int> Replace(this WordprocessingDocument doc,IEnumerable<BatchReplaceStringToTable> replaces)
        {
            var result = new List<int>();
            var elements = doc.FindAllTextElement();
            foreach (var item in replaces)
            {
                result.Add( ReplaceTable(elements,item.OldString,item.Table,item.SaveFormatting));
            }

            return result;
        }

        private static int ReplaceTable(Text[]? elements,string oldString,Table table,bool saveFormatting = true)
        {
            int count = 0;
            var waitReplace = DocElementUtils.FindRun(elements, oldString);

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
                    if (r.RunProperties != null)
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
                if (copy != null)
                {
                    OpenXmlElement element = item[0].Parent;

                    // Table不能随意插入，父元素只能是以下元素
                    while (!tableParentNode.Contains(element.Parent.XName.LocalName))
                    {
                        element = element.Parent;
                    }



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
            return count;
        }

        


        

    }

    /// <summary>
    /// 替换参数组
    /// </summary>
    public record BatchReplaceStringToTable
    {
        /// <summary>
        /// 被替换的字符串
        /// </summary>
        public string OldString { get; set; } = string.Empty;
        /// <summary>
        /// 目标表格
        /// </summary>
        public Table Table { get; set; } = null!;

        /// <summary>
        /// 是否保留原格式
        /// </summary>
        public bool SaveFormatting { get; set; } = true;
    }
}
