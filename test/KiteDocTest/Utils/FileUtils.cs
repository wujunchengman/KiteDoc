using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace KiteDocTest.Utils
{
    public static class FileUtils
    {
        public static string CopyTestFile(string filename)
        {
            var filePath = "StaticResource" + Path.DirectorySeparatorChar + filename;
            var testPath = "StaticResource" + Path.DirectorySeparatorChar + "test" + filename;
            File.Copy(filePath, testPath, true);
            return testPath;
        }
    }
}
