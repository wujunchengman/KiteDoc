using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.ElementBuilder;
using KiteDoc.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KiteDoc
{
    /// <summary>
    /// 通过文本替换元素列表
    /// </summary>
    public static class DocReplaceManyElementExtension
    {
        /// <summary>
        /// 替换Word中的字符串为段落列表
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="oldString">被替换的字符串</param>
        /// <param name="paragraphs">段落列表</param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc, string oldString, List<Paragraph> paragraphs)
        {
            var elements = doc.FindAllTextElement();
            var waitReplace = DocElementUtils.FindRun(elements, oldString);
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
                        el.InsertBeforeSelf<Paragraph>((Paragraph)paragraph.Clone());
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
        /// 替换Word中的字符串为表格列表
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="oldString">被替换的字符串</param>
        /// <param name="tables">表格列表</param>
        /// <param name="appendWhiteParagraph">插入空白段落作为间隔</param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc, string oldString, List<Table> tables, bool appendWhiteParagraph = false)
        {
            var elements = doc.FindAllTextElement();
            var waitReplace = DocElementUtils.FindRun(elements, oldString);
            var count = 0;
            foreach (var item in waitReplace)
            {
                var findFlag = true;
                var el = item.First().Parent;
                while (el is not Paragraph)
                {
                    el = el.Parent;
                    if (el is Body)
                    {
                        findFlag = false;
                        break;
                    }
                }

                if (findFlag)
                {


                    if (appendWhiteParagraph && tables.Count > 1)
                    {
                        var pBuilder = new ParagraphBuilder();
                        pBuilder.AppendText(string.Empty);
                        var p = pBuilder.Build();

                        el.InsertBeforeSelf<Table>((Table)tables.First().Clone());

                        foreach (var table in tables.Skip(1))
                        {
                            el.InsertBeforeSelf((Paragraph)p.Clone());
                            el.InsertBeforeSelf<Table>((Table)table.Clone());
                        }



                    }
                    else
                    {
                        foreach (var table in tables)
                        {

                            el.InsertBeforeSelf<Table>((Table)table.Clone());
                        }
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
        /// 替换Word中的字符串为一组文档元素
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="oldString">被替换的字符串</param>
        /// <param name="openXmlCompositeElements">Word元素对象数据</param>
        /// <returns></returns>
        public static int Replace(this WordprocessingDocument doc, string oldString, List<OpenXmlCompositeElement> openXmlCompositeElements)
        {
            var elements = doc.FindAllTextElement();
            var waitReplace = DocElementUtils.FindRun(elements, oldString);
            var count = 0;
            foreach (var item in waitReplace)
            {
                var findFlag = true;
                var el = item.First().Parent;
                while (el is not Paragraph)
                {
                    el = el.Parent;
                    if (el is Body)
                    {
                        findFlag = false;
                        break;
                    }
                }

                if (findFlag)
                {


                    foreach (var element in openXmlCompositeElements)
                    {

                        el.InsertBeforeSelf((OpenXmlCompositeElement)element.Clone());
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

    }
}
