using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.Enum;
using System.Collections.Generic;

namespace KiteDoc.Interface
{
    /// <summary>
    /// Doc操作辅助接口
    /// </summary>
    public interface IDocOperation
    {
        /// <summary>
        /// 通过书签名查找书签对象
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="bookmarkName">书签名</param>
        /// <returns></returns>
        BookmarkStart FindBookmarkStart(WordprocessingDocument doc, string bookmarkName);

        /// <summary>
        /// 查找书签结束标记
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="id">每一个书签开始会带有一个Id，其值与书签结束所带Id相同</param>
        /// <returns></returns>
        BookmarkEnd FindBookmarkEnd(WordprocessingDocument doc, string id);

        /// <summary>
        /// 删除书签处的内容
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="bookmarkName">书签名</param>
        BookmarkStart RemoveBookmarkContent(WordprocessingDocument doc, string bookmarkName);

        /// <summary>
        /// 删除书签区域（书签所在的段落全部删除）
        /// </summary>
        /// <param name="doc">文档对象</param>
        /// <param name="bookmarkName">书签名</param>
        void RemoveBookmarkArea(WordprocessingDocument doc, string bookmarkName);

        /// <summary>
        /// 向书签处插入文本
        /// </summary>
        /// <param name="bookmarkStart">书签开始对象</param>
        /// <param name="text">文本内容</param>
        /// <param name="fontSize">指定字体大小，不指定时使用默认大小</param>
        void InsertTextToBookmark(BookmarkStart bookmarkStart, string text,float? fontSize);

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
        void ReplaceNormalTableByBookmark(WordprocessingDocument doc, string bookmark, List<string> tableHead,
            List<string[]> data, List<int> widthList, float? fontSize = null, List<HorizontalAlign> aligns = null, bool isSerialNumber = false);


        /// <summary>
        /// 替换书签位置为一个run
        /// </summary>
        /// <param name="doc">word文档对象</param>
        /// <param name="bookmark">书签名</param>
        /// <param name="run">OpenXML的Run对象</param>
        void ReplaceRunByBookmark(WordprocessingDocument doc, string bookmark, Run run);


        /// <summary>
        /// 替换文本内容
        /// </summary>
        /// <param name="doc">文档对象</param>
        /// <param name="originText">原始文本</param>
        /// <param name="destText">目标文本</param>
        void ReplaceText(WordprocessingDocument doc, string originText, string destText);


        /// <summary>
        /// 修改书签文本
        /// </summary>
        /// <param name="doc">word文档对象</param>
        /// <param name="bookmarkName">书签名字</param>
        /// <param name="text">替换的文本</param>
        /// <param name="fontSize">指定字体大小，不指定时使用默认大小</param>
        void ReplaceTextByBookmark(WordprocessingDocument doc, string bookmarkName, string text,float? fontSize=null);

        /// <summary>
        /// 移除文档中所有的页眉页脚
        /// </summary>
        /// <param name="doc">word文档对象</param>
        void RemoveHeadersAndFooters(WordprocessingDocument doc);

        /// <summary>
        /// 将书签区域替换为传入段落列表的内容
        /// 注意：会移除书签所在段落
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="bookmark">书签名</param>
        /// <param name="paragraphs">段落列表</param>
        void ReplaceAreaByBookmark(WordprocessingDocument doc, string bookmark, List<Paragraph> paragraphs);

        /// <summary>
        /// 修改书签处的图片
        /// </summary>
        /// <param name="doc">文档路径</param>
        /// <param name="picPath">图片路径</param>
        /// <param name="imageType">图片类型</param>
        /// <param name="bookmarkName">书签名</param>
        /// <param name="width">图片宽度</param>
        /// <param name="height">图片高度</param>
        void ReplacePictureByBookmark(WordprocessingDocument doc, string picPath, ImageType imageType,
            string bookmarkName, int width = 18,
            int height = -1);
    }
}