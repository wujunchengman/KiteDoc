namespace KiteDoc.Enum
{
    /// <summary>
    /// 单元格合并
    /// </summary>
    public enum TableCellMergeEnum
    {
        /// <summary>
        /// 普通
        /// </summary>
        Normal,
        /// <summary>
        /// 水平合并起始
        /// </summary>
        HorizontalStart,
        /// <summary>
        /// 水平合并单元格(该单元格不会被显示)
        /// </summary>
        HorizontalContinue,
        /// <summary>
        /// 垂直合并单元格起始
        /// </summary>
        VerticalStart,
        /// <summary>
        /// 垂直合并单元格(该单元格不会被显示)
        /// </summary>
        VerticalContinue,
        /// <summary>
        /// 水平合并和垂直合并共同的起始
        /// </summary>
        HorizontalAndVerticalStart,
        /// <summary>
        /// 水平合并和垂直合并共同的合并单元格(该单元格不会被显示)
        /// </summary>
        HorizontalAndVerticalContinue
    }
}