using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using KiteDoc;
using KiteDocTest.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace KiteDocTest
{
    public class PictureTest
    {
        [Fact]
        public void ReplaceMultiplePictureTest()
        {
            var filename = "替换文本为图片.docx";
            string testPath = FileUtils.CopyTestFile(filename, "替换多张图片.docx");

            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var files = System.IO.Directory.GetFiles("C:\\VerificationSystemTestStaticFile\\Pictures");

                var runs = files.Take(6).Select(x => doc.GetPictureRun(x, KiteDoc.Enum.ImageType.Png, 8.85));

                var p = new Paragraph();
                foreach (var item in runs)
                {
                    p.AppendChild(item);
                }

                doc.Replace("假装", new List<Paragraph>() { p });
                //var count = doc.Replace("假装", "D:/CNAS.png", ImageType.Png);
                //Assert.Equal(2, count);
            }
        }
    }
}