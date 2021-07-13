using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc.Enum;
using System.Collections.Generic;

namespace KiteDoc.Interface
{
    public interface IDocElementProvider
    {
        /// <summary>
        /// 返回一个BookmarkStart对象，每一个BookmarkStart必须要有Id相同的BookmarkEnd与之对应
        /// </summary>
        /// <param name="bookmarkId">Bookmark(书签)的唯一标识，用于绑定BookmarkEnd(书签结束)</param>
        /// <param name="bookmarkName">书签名</param>
        /// <returns></returns>
        BookmarkStart GetBookmarkStartObject(int bookmarkId, string bookmarkName);

        /// <summary>
        /// 返回一个BookmarkEnd对象，其Id必须与BookmarkStart的Id相同才能表示一个完整书签
        /// </summary>
        /// <param name="bookmarkId">Bookmark(书签)的唯一标识，必须与BookmarkStart的Id相同才能表示一个完整书签</param>
        /// <returns></returns>
        BookmarkEnd GetBookmarkEndObject(int bookmarkId);

        /// <summary>
        /// 返回一个Run对象
        /// </summary>
        /// <param name="isBold">是否加粗</param>
        /// <param name="fontSize">字体大小，传递为Null则使用默认字体大小</param>
        /// <param name="font">字体，默认为宋体</param>
        /// <returns></returns>
        Run GetRunObject(bool isBold, float? fontSize = null, string font = "宋体");

        /// <summary>
        /// 返回一个Paragraph段落对象
        /// </summary>
        /// <param name="align">对齐方式，默认为左对齐</param>
        /// <param name="indentationForTheFirstLine">首行缩进，默认为否</param>
        /// <returns></returns>
        Paragraph GetParagraphObject(
            HorizontalAlign align = HorizontalAlign.Left,
            bool indentationForTheFirstLine = false);

        /// <summary>
        /// 返回一个单元格对象
        /// </summary>
        /// <param name="width">宽度百分比</param>
        /// <param name="verticalAlign">垂直对齐方式，默认居中对齐</param>
        /// <param name="cellMerge">单元格合并，默认为普通，即不参与合并</param>
        /// <returns></returns>
        TableCell GetTableCellObject(
            int width,
            VerticalAlign verticalAlign = VerticalAlign.Center,
            DocCellMergeEnum cellMerge = DocCellMergeEnum.Normal
        );

        /// <summary>
        /// 返回一个表格对象
        /// </summary>
        /// <param name="isBorder">是否启用表格边框，默认启用</param>
        /// <param name="width">表格宽度(百分比),默认100</param>
        /// <returns></returns>
        Table GetTableObject(bool isBorder = true, int width = 100);

        /// <summary>
        /// 生成表格，如果单元格没数据会自动合并单元格
        /// </summary>
        /// <param name="tableHead">表头，有空行添加空数据</param>
        /// <param name="data">表格数据</param>
        /// <param name="widthList">表格宽度列表（单位：%）</param>
        /// <param name="fontSize">字号，为null时使用默认字号</param>
        /// <param name="align">对齐方式，默认居中对齐</param>
        /// <param name="isSerialNumber">单元格内换行数据是否编号</param>
        /// <returns></returns>
        Table GenerateNormalTable(List<string> tableHead, List<string[]> data, List<int> widthList,
            float? fontSize, HorizontalAlign align = HorizontalAlign.Center, bool isSerialNumber = false);

        /// <summary>
        /// 生成表格，如果单元格没数据会自动合并单元格
        /// </summary>
        /// <param name="tableHead">表头，有空行添加空数据</param>
        /// <param name="data">表格数据</param>
        /// <param name="widthList">表格宽度列表（单位：%）</param>
        /// <param name="fontSize">字体大小，默认为null，为null时</param>
        /// <param name="align">全局对齐方式，默认居中对齐</param>
        /// <param name="isSerialNumber">单元格内换行数据是否编号</param>
        /// <returns></returns>
        Table GenerateNormalTable(List<string> tableHead, List<string[]> data, List<int> widthList,
            float? fontSize, List<HorizontalAlign> align, bool isSerialNumber = false);

        /// <summary>
        /// 获得一个包含垂直合并的表格
        /// 如果为null则进行垂直合并
        /// </summary>
        /// <param name="tableHead">表头</param>
        /// <param name="data">表格数据</param>
        /// <param name="widthList">表格宽度列表，不提供时为100%均分</param>
        /// <param name="fontSize">字体大小，默认为null，为null时继承上一级字体大小</param>
        /// <param name="align">对齐方式，默认左对齐</param>
        /// <param name="isSerialNumber">单元格内换行数据是否编号</param>
        /// <returns></returns>
        Table GenerateVerticalMergeTable(List<string> tableHead, List<string[]> data, List<int> widthList,
            float? fontSize = null, List<HorizontalAlign> align = null, bool isSerialNumber = false);

        /// <summary>
        /// 获得一个包含图片的Run对象
        /// </summary>
        /// <param name="doc">Word文档对象</param>
        /// <param name="fileName">图片文件名（全路径）</param>
        /// <param name="imageType">图片类型</param>
        /// <param name="width">图片宽度，单位cm</param>
        /// <param name="height">图片高度，单位cm</param>
        Run GetPictureRun(WordprocessingDocument doc, string fileName, ImageType imageType, int width = 18,
            int height = -1);
    }
}