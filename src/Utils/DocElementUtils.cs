using DocumentFormat.OpenXml.Packaging;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;

namespace KiteDoc.Utils
{
    public static class DocElementUtils
    {
        public static Text[] FindAllTextElement(this WordprocessingDocument doc)
        {
            if (doc.MainDocumentPart != null)
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

        public static List<List<Run>> FindRun(Text[]? elements, string oldString)
        {
            var waitReplace = new List<List<Run>>();
            var continueFlag = false;
            var tempString = new StringBuilder();
            var tempRuns = new Dictionary<int, Run>();
            var splitRuns = new Dictionary<int, Run>();

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

                    // 一定有匹配，不确定是不是部分匹配(一旦有字符不匹配就会设为-1，所以不为-1时有可能是长度不够，也有可能是完全匹配)
                    var index = thisTextString.PartContains(oldString);
                    if (index != -1)
                    {
                        // 完全匹配
                        if (thisTextString == oldString)
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
                            splitRuns.Add(i, (Run)elements[i].Parent);
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
                            if (innerIndex > 0)
                            {
                                // 拼接对应的文本放入缓存中
                                tempString.Append(thisTextString);
                                continueFlag = true;
                                // 要替换的Run添加进去，因为是内包含，所以一定是要替换的
                                tempRuns.Add(i, (Run)elements[i].Parent);
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
                            waitReplace.Add(tempRuns.OrderBy(x => x.Key).Select(x => x.Value).ToList());
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
                                var runs = tempRuns.Concat(splitRuns).OrderBy(x => x.Key).Select(x => x.Value).ToList();
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
    }
}