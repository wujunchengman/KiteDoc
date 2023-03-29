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
            var elements = doc.MainDocumentPart.Document.Descendants<Text>().ToArray();
            int count = 0;
            var waitReplace = FindRun(elements, oldString);
            foreach (var item in waitReplace)
            {
                if (item.Count==1)
                {
                    var text = item[0].Descendants<Text>().FirstOrDefault();
                    text.Text = text.Text.Replace(oldString, newString);
                }
                else
                {
                    var text = item[0].Descendants<Text>().FirstOrDefault();

                    // 把所有Run的Text拼到一起，然后去替换对应的文本
                    var dest = string.Concat(item.Select(x => x.Descendants<Text>().FirstOrDefault()).Select(t => t.Text));

                    text.Text = dest.Replace(oldString,newString);

                    // todo: 需要考虑最后一个Run对象中是否还有其他的文本
                    foreach (var run in item.Skip(1))
                    {
                        run.Remove();
                    }
                }
                count++;

            }


            
            //foreach (var text in elements)
            //{
            //    if (text.Text.Contains(oldString))
            //    {
            //        text.Text = text.Text.Replace(oldString, newString);
            //        count++;
            //    }
            //}
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

            var waitReplace = FindRun(elements,oldString);

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
                            waitReplace.Add(tempRuns.Select(x=>x.Value).ToList());
                        }
                        // 因为引用传递的关系，这里不能直接clear
                        tempRuns = new Dictionary<int, Run>();

                        continueFlag = false;


                    }
                    else
                    {
                        // 符合的情况

                        // 完全匹配
                        if (tempStringValue.Length == oldString.Length)
                        {
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
                        else if (index + oldString.Length > thisTextString.Length)
                        {
                            // 拼接对应的文本放入缓存中
                            continueFlag = true;
                            // 要替换的Run添加进去
                            splitRuns.Add(i, (Run)elements[i].Parent);
                        }
                        // 内包含
                        else
                        {
                            var innerIndex = thisTextString[(thisTextString.Length - oldString.Length)..].PartContains(oldString);
                            // 内包含+尾部重叠
                            if (innerIndex != -1)
                            {
                                continueFlag = true;
                                // 要替换的Run添加进去，因为是内包含，所以一定是要替换的
                                tempRuns.Add(i, (Run)elements[i].Parent);
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





                //tempString.Append(elements[i].Text);
                //// 获取当前Text对象的父对象Run
                //tempRuns.Add((Run)elements[i].Parent);

                //// todo: 查找时，应该先判断是否存在包含，移除了包含后的字符串，是否与开头匹配

                //// 判断是否是包含关系
                
                
                //var indexOf = thisTextString.IndexOf(oldString);
                //// 存在包含关系
                //if (indexOf!=-1)
                //{
                //    thisTextString = thisTextString[(indexOf + oldString.Length)..];
                //}
                //else
                //{
                //    // 查找是否存在部分匹配的情况
                //}


                //if (!continueFlag && thisTextString.Contains(oldString))
                //{
                //    thisTextString.Substring(1);
                //}

                //// 如果第一次找到开头
                //if (!continueFlag && oldString.StartsWith(elements[i].Text))
                //{
                //    continueFlag = true;
                //}

                //// 如果添加了之后还是满足开始，说明还可以考虑向后加
                //else if (continueFlag &&
                //    oldString.StartsWith(tempString.ToString()) &&
                //    // 只有所有Run的父元素相同的情况下才可以考虑合并文本信息
                //    elements[i].Parent.Parent == tempRuns[^1].Parent
                //    )
                //{
                //}
                //// 追加了之后不满足了，说明只有前面一部分相同，不是完全相同
                //else if (continueFlag)
                //{
                //    // 将继续标识置为False，下次重新匹配
                //    continueFlag = false;

                //    tempString.Clear();
                //    tempRuns.Clear();
                //}
                //else
                //{
                //    tempString.Clear();
                //    tempRuns.Clear();
                //}


                //// 如果匹配成功了，则添加到待替换列表，同时重置待匹配列表和标记
                //if (continueFlag && tempString.ToString().Contains(oldString))
                //{
                //    continueFlag = false;
                //    waitReplace.Add(tempRuns);

                //    // 因为后面会添加到数组中，考虑引用传递的问题，这里重新new一个数组
                //    tempRuns = new List<Run>();
                //    // 清空比对字符串
                //    tempString.Clear();
                //}
            }

            return waitReplace;
        }


    }
}
