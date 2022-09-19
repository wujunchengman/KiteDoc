using System;
using System.Collections.Generic;
using System.IO;
using DocumentFormat.OpenXml.Packaging;
using KiteDoc;
using KiteDoc.Interface;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var filePath = "D:\\WorkFiles\\VerificationSystemStaticFile\\测试文档.docx";
            File.Copy(filePath, filePath.Replace("测试文档", "测试文档2"),true);
            filePath = "D:\\WorkFiles\\VerificationSystemStaticFile\\测试文档2.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath,true ))
            {
                IDocOperation operation = new DocOperation(new DocElementProvider());
                operation.ReplaceText(doc,"假装","江南皮革厂");

                //operation.ReplaceTextByBookmark(doc, "test", "通过书签替换的内容");
            }
            Console.WriteLine("Hello World!");
        }
    }
}
