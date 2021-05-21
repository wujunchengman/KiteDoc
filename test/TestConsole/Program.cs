using System;
using System.Collections.Generic;
using KiteDoc;
using KiteDoc.Interface;

namespace TestConsole
{
    class Program
    {
        static void Main(string[] args)
        {
            var analysisData = new List<string[]>()
            {
                new[]
                {
                    "测点终端安装数量","1个"
                },
                new[]
                {
                    "测点终端安装位置","测点ch5附近"
                },
                new[]
                {
                    "测点终端温度测量范围","（-40~60）℃"
                },
                new []
                {
                    "测点终端报警功能",
                    "温度≤" + 3 + "℃：声光报警"
                },
                new []
                {
                    null,
                    "（" + 3 +  "～" +  7 + "）℃：声光报警"
                },
                new []
                {
                    null,
                    "温度≥" + 7 + "℃：声光报警"
                },
            };

            IDocElementProvider docElementProvider = new DocElementProvider();

            var widthList = new List<int>()
             {
                 50, 50
             };

            docElementProvider.GenerateVerticalMergeTable(null, analysisData, widthList, null);
            Console.WriteLine("Hello World!");
        }
    }
}
