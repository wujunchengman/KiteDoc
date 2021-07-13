using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.Enum;
using KiteDoc.Interface;
using System;
using System.Collections.Generic;
using System.Linq;

namespace KiteDoc
{
    public class DocOperation : IDocOperation
    {
        private readonly IDocElementProvider _docElementProvider;

        public DocOperation(IDocElementProvider docElementProvider)
        {
            _docElementProvider = docElementProvider;
        }
        /// <summary>
        /// 通过书签名查找书签对象
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="bookmarkName">书签名</param>
        /// <returns></returns>
        public BookmarkStart FindBookmarkStart(WordprocessingDocument doc, string bookmarkName)
        {
            // 在正文中查找书签开始标记
            foreach (var inst in doc.MainDocumentPart.RootElement.Descendants<BookmarkStart>())
            {
                if (inst.Name == bookmarkName)
                {
                    return inst;
                }
            }

            // 在页眉查找书签开始标记
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                foreach (var inst in header.Header.Descendants<BookmarkStart>())
                {
                    if (inst.Name == bookmarkName)
                    {
                        return inst;
                    }
                }
            }

            // 在页脚查找书签开始标记
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                foreach (var inst in footer.Footer.Descendants<BookmarkStart>())
                {
                    if (inst.Name == bookmarkName)
                    {
                        return inst;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// 查找书签结束标记
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="id">每一个书签开始会带有一个Id，其值与书签结束所带Id相同</param>
        /// <returns></returns>
        public BookmarkEnd FindBookmarkEnd(WordprocessingDocument doc, string id)
        {
            // 在正文中查找书签结束标记
            foreach (var inst in doc.MainDocumentPart.RootElement.Descendants<BookmarkEnd>())
            {
                if (inst.Id == id)
                {
                    return inst;
                }
            }

            // 在页脚查询书签结束标记
            foreach (var footer in doc.MainDocumentPart.FooterParts)
            {
                foreach (var inst in footer.Footer.Descendants<BookmarkEnd>())
                {
                    if (inst.Id == id)
                    {
                        return inst;
                    }
                }
            }

            // 在页眉查找书签结束标记
            foreach (var header in doc.MainDocumentPart.HeaderParts)
            {
                foreach (var inst in header.Header.Descendants<BookmarkEnd>())
                {
                    if (inst.Id == id)
                    {
                        return inst;
                    }
                }
            }

            // 如果都没查到则返回null
            return null;
        }

        /// <summary>
        /// 删除书签处的内容
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="bookmarkName">书签名</param>
        public BookmarkStart RemoveBookmarkContent(WordprocessingDocument doc, string bookmarkName)
        {
            BookmarkStart bookmarkStart = FindBookmarkStart(doc, bookmarkName);
            BookmarkEnd bookmarkEnd = FindBookmarkEnd(doc, bookmarkStart.Id);
            while (true)
            {
                // NextSibling()获取紧跟当前 OpenXmlElement 元素的 OpenXmlElement 元素。 如果没有下一个 OpenXmlElement 元素，则返回 null 
                var run = bookmarkStart.NextSibling();
                if (run == null)
                {
                    break;
                }

                if (run is BookmarkEnd end && end == bookmarkEnd)
                {
                    break;
                }

                run.Remove();
            }

            return bookmarkStart;
        }

        /// <summary>
        /// 删除书签区域（书签所在的段落全部删除）
        /// </summary>
        /// <param name="doc">文档对象</param>
        /// <param name="bookmarkName">书签名</param>
        public void RemoveBookmarkArea(WordprocessingDocument doc, string bookmarkName)
        {
            BookmarkStart bookmarkStart = FindBookmarkStart(doc, bookmarkName);
            BookmarkEnd bookmarkEnd = FindBookmarkEnd(doc, bookmarkStart.Id);

            var bookmarkFather = bookmarkStart.Parent;
            var bookmarkEndFather = bookmarkEnd.Parent;
            if (bookmarkFather != bookmarkEndFather)
            {
                while (true)
                {
                    var element = bookmarkFather.NextSibling<Paragraph>();
                    if (element == bookmarkEndFather)
                    {
                        break;
                    }

                    if (element == null)
                    {
                        break;
                    }

                    if (element is Paragraph)
                    {
                        element.Remove();
                    }
                }
            }

            bookmarkFather.Remove();
        }

        /// <summary>
        /// 向书签处插入文本
        /// </summary>
        /// <param name="bookmarkStart">书签开始对象</param>
        /// <param name="text">文本内容</param>
        /// <param name="fontSize">指定字体大小，不指定时使用默认大小</param>
        public void InsertTextToBookmark(BookmarkStart bookmarkStart, string text,float? fontSize=null)
        {
            // Run run = new Run(new Text(text));
            var run = _docElementProvider.GetRunObject(false,fontSize);
            run.AppendChild(new Text(text));
            bookmarkStart.Parent.InsertAfter<Run>(run, bookmarkStart);
        }



        /// <summary>
        /// 替换书签位置为普通表格
        /// </summary>
        /// <param name="doc">word文档对象</param>
        /// <param name="bookmark">书签名</param>
        /// <param name="tableHead">表头</param>
        /// <param name="data">表格数据，通过#$$#分隔，即通过分隔一个单元格会有多个段落</param>
        /// <param name="widthList">宽度百分比列表</param>
        /// <param name="fontSize">字体大小，默认为null，为null时</param>
        /// <param name="aligns">水平对齐列表，数量等于列数时为每一列单独指定，只有一个时为全局指定</param>
        /// <param name="isSerialNumber">单元格内换行数据是否编号</param>
        public void ReplaceNormalTableByBookmark(WordprocessingDocument doc, string bookmark, List<string> tableHead,
            List<string[]> data, List<int> widthList, float? fontSize = null, List<HorizontalAlign> aligns = null, bool isSerialNumber = false)
        {
            RemoveBookmarkContent(doc, bookmark);
            BookmarkStart bookmarkStart = FindBookmarkStart(doc, bookmark);
            if (bookmarkStart == null)
            {
                return;
            }

            Run run = null;
            if (aligns == null)
            {
                run = new Run(new Table(_docElementProvider.GenerateNormalTable(tableHead, data, widthList, fontSize)));

            }
            else if (aligns.Count == 1)
            {
                run = new Run(new Table(_docElementProvider.GenerateNormalTable(tableHead, data, widthList, fontSize, aligns[0], isSerialNumber)));
            }
            else if (aligns.Count == data[0].Length)
            {
                run = new Run(new Table(_docElementProvider.GenerateNormalTable(tableHead, data, widthList, fontSize, aligns, isSerialNumber)));
            }

            bookmarkStart.Parent.InsertAfter<Run>(run, bookmarkStart);
        }


        /// <summary>
        /// 替换书签位置为一个run
        /// </summary>
        /// <param name="doc">word文档对象</param>
        /// <param name="bookmark">书签名</param>
        /// <param name="run">OpenXML的Run对象</param>
        public void ReplaceRunByBookmark(WordprocessingDocument doc, string bookmark, Run run)
        {
            RemoveBookmarkContent(doc, bookmark);
            BookmarkStart bookMarkStart = FindBookmarkStart(doc, bookmark);
            if (bookMarkStart == null)
            {
                return;
            }
            bookMarkStart.Parent.InsertAfter<Run>(run, bookMarkStart);
        }

        /// <summary>
        /// 移除文档中所有的页眉页脚
        /// </summary>
        /// <param name="doc">word文档对象</param>
        public void RemoveHeadersAndFooters(WordprocessingDocument doc)
        {
            // Get a reference to the main document part.
            var docPart = doc.MainDocumentPart;

            // Count the header and footer parts and continue if there 
            // are any.
            if (docPart.HeaderParts.Count() > 0 ||
                docPart.FooterParts.Count() > 0)
            {
                // Remove the header and footer parts.
                docPart.DeleteParts(docPart.HeaderParts);
                docPart.DeleteParts(docPart.FooterParts);

                // Get a reference to the root element of the main
                // document part.
                Document document = docPart.Document;

                // Remove all references to the headers and footers.

                // First, create a list of all descendants of type
                // HeaderReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var headers =
                  document.Descendants<HeaderReference>().ToList();
                foreach (var header in headers)
                {
                    header.Remove();
                }

                // First, create a list of all descendants of type
                // FooterReference. Then, navigate the list and call
                // Remove on each item to delete the reference.
                var footers =
                  document.Descendants<FooterReference>().ToList();
                foreach (var footer in footers)
                {
                    footer.Remove();
                }

                // Save the changes.
                document.Save();
            }
        }

        /// <summary>
        /// 将书签区域替换为传入段落列表的内容
        /// 注意：会移除书签所在段落
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="paragraphs">段落列表</param>
        public void ReplaceAreaByBookmark(WordprocessingDocument doc, string bookmarkName, List<Paragraph> paragraphs)
        {
            BookmarkStart bookmarkStart = FindBookmarkStart(doc, bookmarkName);
            BookmarkEnd bookmarkEnd = FindBookmarkEnd(doc, bookmarkStart.Id);

            var bookmarkFather = bookmarkStart.Parent;
            var bookmarkEndFather = bookmarkEnd.Parent;
            if (bookmarkFather != bookmarkEndFather)
            {
                while (true)
                {
                    var element = bookmarkFather.NextSibling();
                    if (element == bookmarkEndFather || element == bookmarkEnd)
                    {
                        break;
                    }

                    if (element == null)
                    {
                        break;
                    }


                    element.Remove();
                }
            }

            InsertParagraphToBookmark(doc, bookmarkStart, paragraphs);
            // 移除书签所在段落，避免出现插入的内容后面多一个回车段落的情况
            bookmarkFather.Remove();
            // RemoveBookmarkArea(doc,bookmark);
        }

        /// <summary>
        /// 向书签处插入一个段落
        /// </summary>
        /// <param name="doc">文档对象</param>
        /// <param name="bookmarkStart">书签开始对象</param>
        /// <param name="paragraphs">段落列表对象</param>
        public void InsertParagraphToBookmark(WordprocessingDocument doc, BookmarkStart bookmarkStart,
        List<Paragraph> paragraphs)
        {
            if (bookmarkStart == null)
            {
                throw new Exception("传递了空的书签开始对象");
            }

            // 在书签位置的后面插入，其实后一个元素是在上一个元素前面，所以将Paragraph数组反转一下
            paragraphs.Reverse();

            // 插入段落对象
            foreach (var paragraph in paragraphs)
            {
                var p = bookmarkStart.Parent;

                p.InsertAfterSelf(paragraph.CloneNode(true));


            }

            // 将反转的Paragraph数组还原，后面可能会继续用
            paragraphs.Reverse();
        }

        /// <summary>
        /// 替换文本内容
        /// </summary>
        /// <param name="doc">文档对象</param>
        /// <param name="originText">原始文本</param>
        /// <param name="destText">目标文本</param>
        public void ReplaceText(WordprocessingDocument doc,string originText,string destText)
        {
            // 替换正文中的内容
            var body = doc.MainDocumentPart.Document.Body;
            {
                var paras = body.Elements<Paragraph>();
                foreach (var para in paras)
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            if (text.Text.Contains(originText))
                            {
                                text.Text = text.Text.Replace(originText, destText);
                            }
                        }
                    }
                }
            }
            // 替换页脚的内容
            var footer = doc.MainDocumentPart.FooterParts;
            foreach (var footerPart in footer)
            {
                var paras = footerPart.Footer.Elements<Paragraph>();
                foreach (var para in paras)
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
                        {
                            if (text.Text.Contains(originText))
                            {
                                text.Text = text.Text.Replace(originText, destText);
                            }
                        }
                    }
                }
            }

            var header = doc.MainDocumentPart.HeaderParts;
            foreach (var headerPart in header)
            {
                var paras = headerPart.Header.Elements<Paragraph>();
                foreach (var para in paras)
                {
                    foreach (var run in para.Elements<Run>())
                    {
                        foreach (var text in run.Elements<Text>())
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

        /// <summary>
        /// 修改书签文本
        /// </summary>
        /// <param name="doc">word文档对象</param>
        /// <param name="bookmarkName">书签名字</param>
        /// <param name="text">替换的文本</param>
        /// <param name="fontSize">指定字体大小，不指定时使用默认大小</param>
        public void ReplaceTextByBookmark(WordprocessingDocument doc, string bookmarkName, string text,float? fontSize = null)
        {
            var bookmarkStart = RemoveBookmarkContent(doc, bookmarkName);
            InsertTextToBookmark(bookmarkStart, text,fontSize);
        }

        /// <summary>
        /// 修改书签处的图片
        /// </summary>
        /// <param name="doc">文档路径</param>
        /// <param name="picPath">图片路径</param>
        /// <param name="imageType">图片类型</param>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="width">图片宽度</param>
        /// <param name="height">图片高度</param>
        public void ReplacePictureByBookmark(WordprocessingDocument doc, string picPath, ImageType imageType, string bookmarkName, int width = 18,
            int height = -1)
        {
            var bookmarkStart = RemoveBookmarkContent(doc, bookmarkName);
            var run = _docElementProvider.GetPictureRun(doc, picPath, imageType, width, height);
            bookmarkStart.Parent.InsertAfter(run, bookmarkStart);
        }

    }
}