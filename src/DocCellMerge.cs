namespace KiteDoc
{
    public enum CellMergeEnum
    {
        Horizontal,
        Vertical
    }

    public class DocCellMerge
    {
        /// <summary>
        /// 合并开始索引
        /// </summary>
        public int Start { get; set; }
        /// <summary>
        /// 合并结束索引
        /// </summary>
        public int Finish { get; set; }
        /// <summary>
        /// 定位，标识合并所在的行或列
        /// </summary>
        public int Position { get; set; }
        /// <summary>
        /// 合并方式
        /// </summary>
        public CellMergeEnum CellMerge { get; set; }
    }
}