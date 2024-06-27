using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using KiteDoc.Enum;
using KiteDoc.Utils;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;

namespace KiteDoc
{
    /// <summary>
    /// 替换文本为图片
    /// </summary>
    public static class DocReplacePictureExtension
    {
        public static int Replace(this WordprocessingDocument doc, string oldString, string fileName, ImageType imageType, double width = 18, double height = -1)
        {
            var count = 0;

            var elements = doc.FindAllTextElement();
            var waitReplace = DocElementUtils.FindRun(elements, oldString);

            if (waitReplace.Any())
            {
                var run = doc.GetPictureRun(fileName, imageType, width, height);

                foreach (var item in waitReplace)
                {
                    // 获取图片，插入到对应的位置
                    item.First().InsertAfterSelf((Run)run.Clone());

                    foreach (var waitDelete in item)
                    {
                        waitDelete.Remove();
                    }

                    count++;
                }
            }

            return count;
        }
    }
}