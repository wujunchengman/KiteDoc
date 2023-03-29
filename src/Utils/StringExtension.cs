using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KiteDoc.Utils
{
    /// <summary>
    /// String扩展
    /// </summary>
    public static class StringExtension
    {
        /// <summary>
        /// 查找是否存在部分包含
        /// </summary>
        /// <param name="source"></param>
        /// <param name="match"></param>
        /// <returns></returns>
        public static int PartContains(this string source, string match)
        {
            var index = 0;
            var i = 0;


            while (i < source.Length) {
                if (source[i] == match[index])
                {
                    index++;

                    // 放在break之前，这样索引才能计算正确
                    i++;

                    // 如果到了最长，说明是包含关系
                    if (index == match.Length)
                        break;
                }
                else
                {
                    index = 0;

                    // 没有匹配上进一匹配下一个
                    i++;
                }
            };


            return index>0?i-index:-1;
        }
    }
}
