using System;
using System.Collections.Generic;
using DocumentFormat.OpenXml.Packaging;
using KiteDoc;
using KiteDoc.Interface;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var filePath = "D:\\VerificationSystemStaticFile\\测试文档.docx";
            using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath,true ))
            {
                IDocOperation operation = new DocOperation(new DocElementProvider());
                operation.ReplaceText(doc,"假装","江南皮革厂");
            }
            Console.WriteLine("Hello World!");
        }
    }
}
