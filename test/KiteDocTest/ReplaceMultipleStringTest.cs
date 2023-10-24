using DocumentFormat.OpenXml.Packaging;
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
    public class ReplaceMultipleStringTest
    {
        [Fact]
        public void ReplaceNotExistKeyText()
        {
            var filename = "替换文本.docx";
            string testPath = FileUtils.CopyTestFile(filename);




            using (WordprocessingDocument doc = WordprocessingDocument.Open(testPath, true))
            {
                var dic = new Dictionary<string, string>() {
                    {"假装","改变之后的目标文本" },
                    {"Test","Result" },
                };

                var count = doc.Replace(dic);
                Assert.Equal(new List<int> { 5, 0 }, count);
            }
            //File.Delete(testPath);
        }
    }
}
