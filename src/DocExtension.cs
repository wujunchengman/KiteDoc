using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Drawing.Charts;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.Utils;
using SixLabors.ImageSharp.Formats.Tiff.Compression.Decompressors;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml;
using System.Xml.Linq;

namespace KiteDoc
{
    /// <summary>
    /// 针对OpenXML提供的一系列扩展方法
    /// </summary>
    public static class DocExtension
    {
        /// <summary>
        /// 替换Word中的字符串为段落列表
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="oldString">被替换的字符串</param>
        /// <param name="paragraphs">段落列表</param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc,string oldString,List<Paragraph> paragraphs)
        {
            var elements = doc.FindAllTextElement();
            var waitReplace = FindRun(elements, oldString);
            var count = 0;
            foreach (var item in waitReplace)
            {
                var pFlag = true;
                var el = item.First().Parent;
                while (el is not Paragraph)
                {
                    el = el.Parent;
                    if (el is Body)
                    {
                        pFlag = false;
                        break;
                    }
                }

                if (pFlag)
                {
                    foreach (var paragraph in paragraphs)
                    {
                        el.InsertBeforeSelf<Paragraph>(paragraph.Clone() as Paragraph);
                    }

                    el.Remove();

                    count++;
                }
                else
                {
                    continue;
                }

            }

            return count;
        }


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
            var waitReplace = FindRun(elements, oldString);
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
            var waitReplace = FindRun(elements, oldString);

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

        private static List<List<Run>> FindRun(Text[]? elements,string oldString)
        {
            var waitReplace = new List<List<Run>>();
            var continueFlag = false;
            var tempString = new StringBuilder();
            var tempRuns = new Dictionary<int,Run>();
            var splitRuns = new Dictionary<int,Run>();

            for (int i = 0; i < elements.Length; i++)
            {

                // 如果是从头开始的，而不是因为部分匹配继续的
                if (!continueFlag)
                {
                    // 获取当前Text对象的文本信息
                    var thisTextString = elements[i].Text;

                    /*
                     * 如果存在包含，则当前的run一定要添加
                     * 如果部分匹配，且加上后面run能够匹配，则将后面的run一起添加
                     * 如果有包含有又有部分匹配，则将当前run和后面的run一起添加
                     */

                    // 一定有匹配，不确定是不是部分匹配
                    var index = thisTextString.PartContains(oldString);
                    if (index != -1)
                    {

                        // 完全匹配
                        if (thisTextString.Length == oldString.Length)
                        {
                            // 添加到待替换列表中
                            waitReplace.Add(new List<Run> { (Run)elements[i].Parent });
                        }
                        // 尾部重叠
                        else if (index + oldString.Length > thisTextString.Length)
                        {
                            // 拼接对应的文本放入缓存中
                            tempString.Append(thisTextString);
                            continueFlag = true;
                            // 要替换的Run添加进去
                            splitRuns.Add(i,(Run)elements[i].Parent);
                        }
                        // 内包含
                        else
                        {
                            var innerIndex = thisTextString[(thisTextString.Length - oldString.Length)..].PartContains(oldString);
                            
                            /*
                             * 取值可能
                             * -1 尾部不一样
                             * 0 尾部一模一样
                             * >0 尾部部分重叠
                             */
                            
                            // 内包含+尾部重叠
                            if (innerIndex >0 )
                            {
                                // 拼接对应的文本放入缓存中
                                tempString.Append(thisTextString);
                                continueFlag = true;
                                // 要替换的Run添加进去，因为是内包含，所以一定是要替换的
                                tempRuns.Add(i,(Run)elements[i].Parent);
                            }
                            // 单纯的内包含
                            else
                            {
                                // 添加到待替换列表中
                                waitReplace.Add(new List<Run> { (Run)elements[i].Parent });
                            }
                        }

                    }
                }
                // 从上一次部分匹配添加过来的
                else
                {
                    // 获取当前Text对象的文本信息
                    var thisTextString = elements[i].Text;
                    tempString.Append(thisTextString);
                    var tempStringValue = tempString.ToString();

                    var index = tempStringValue.PartContains(oldString);
                    // 判断是否符合
                    if (index == -1)
                    {
                        // 不符合的情况


                        // 当找到不符合时就提交所有需要提交的
                        splitRuns.Clear();
                        if (tempRuns.Any())
                        {
                            waitReplace.Add(tempRuns.OrderBy(x=>x.Key).Select(x=>x.Value).ToList());
                            tempRuns.Clear();
                        }

                        tempString.Clear();
                        continueFlag = false;


                    }
                    else
                    {
                        // 符合的情况

                        // 完全匹配
                        if (tempStringValue.Length == oldString.Length)
                        {
                            // 添加了本次后完全匹配，则本次的依然在内
                            splitRuns.Add(i, (Run)elements[i].Parent);

                            // 拼接split和temp
                            var runs = tempRuns.Concat(splitRuns).OrderBy(x => x.Key).Select(x => x.Value).ToList();
                            // 添加到待替换列表中
                            waitReplace.Add(runs);
                            splitRuns.Clear();
                            tempRuns.Clear();
                            tempString.Clear();
                            continueFlag = false;
                        }
                        // 尾部重叠
                        else if (index + oldString.Length > tempStringValue.Length)
                        {
                            // 拼接对应的文本放入缓存中
                            continueFlag = true;
                            // 要替换的Run添加进去
                            splitRuns.Add(i, (Run)elements[i].Parent);
                        }
                        // 内包含
                        else
                        {
                            var endTempStringValue = tempStringValue[(tempStringValue.Length - oldString.Length)..];
                            var innerIndex = endTempStringValue.PartContains(oldString);
                            // 内包含+尾部重叠
                            if (innerIndex != -1)
                            {
                                /*
                                 * 考虑和前面的文本构成一个关键字，同时还有可能是后一个关键字的起点，并且也许中间还夹杂着一个完整的关键字
                                 * 重点是不知道是不是后一个关键字的起点，如果是，则要将后续的添加到本次中，如果不是，这些也要提交
                                 */

                                // 要替换的Run添加进去，因为是内包含，所以一定是要替换的
                                splitRuns.Add(i, (Run)elements[i].Parent);

                                // 考虑有可能是后一个尾部的情况，前面的肯定是可以提交了，因此把前面的放在tempRuns这个确定要提交的字典中
                                // 然后将比对文件截断，仅留后面的部分去匹配，如果匹配上了就继续
                                tempRuns.Concat(splitRuns);
                                splitRuns.Clear();
                                tempString.Clear();
                                tempString.Append(endTempStringValue);
                                

                                continueFlag = true;

                            }
                            // 单纯的内包含
                            else
                            {
                                // 拼接split和temp
                                var runs = tempRuns.Concat(splitRuns).OrderBy(x=>x.Key).Select(x=>x.Value).ToList();
                                // 添加到待替换列表中
                                waitReplace.Add(runs);
                                splitRuns.Clear();
                                tempRuns.Clear();
                                tempString.Clear();
                                continueFlag = false;
                            }
                        }

                    }
                }

            }

            return waitReplace;
        }


        private static Text[] FindAllTextElement(this WordprocessingDocument doc)
        {
            if (doc.MainDocumentPart!=null)
            {
                var result = doc.MainDocumentPart.Document.Descendants<Text>();

                foreach (var element in doc.MainDocumentPart.FooterParts)
                {
                    result = result.Concat(element.Footer.Descendants<Text>());
                }

                foreach (var element in doc.MainDocumentPart.HeaderParts)
                {
                    result = result.Concat(element.Header.Descendants<Text>());
                }

                return result.ToArray();
            }
            
            throw new ArgumentNullException(nameof(doc.MainDocumentPart));
            
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
