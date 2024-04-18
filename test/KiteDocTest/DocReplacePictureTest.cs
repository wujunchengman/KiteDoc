using DocumentFormat.OpenXml.Packaging;
using KiteDoc;
using KiteDoc.Enum;
using KiteDocTest.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace KiteDocTest
{
    public class DocReplacePictureTest
    {
        [Fact]
        public void ReplacePictureTest()
        {
            var filename = "替换文本为图片.docx";
            string testPath = FileUtils.CopyTestFile(filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var count = doc.Replace("假装", "D:/CNAS.png", ImageType.Png);
                Assert.Equal(2, count);
            }
        }

        [Fact]
        public void ReplacePicture_2_Test()
        {
            var filename = "保温箱验证方案.docx";
            string testPath = FileUtils.CopyTestFile(filename);

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var count = doc.Replace("{LayoutDrawing}", "D:/CNAS.png", ImageType.Png);
                Assert.Equal(1, count);
            }
        }
    }
}