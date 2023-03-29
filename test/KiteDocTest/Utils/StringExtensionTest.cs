using KiteDoc.Utils;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Xunit;

namespace KiteDocTest.Utils
{
    public class StringExtensionTest
    {
        [Fact]
        public void PartContainsTestStart()
        {
            var source = "abcdefg";
            var startIndex = source.PartContains("abcde");

            Assert.Equal(0, startIndex);
        }

        [Fact]
        public void PartContainsTestMiddle()
        {
            var source = "abcdefg";
            var startIndex = source.PartContains("ef");

            Assert.Equal(4, startIndex);
        }


        [Fact]
        public void PartContainsTestNotContains()
        {
            var source = "abcdefg";
            var startIndex = source.PartContains("efd");

            Assert.Equal(-1, startIndex);
        }


        [Fact]
        public void PartContainsTestPartContains()
        {
            var source = "abcdefg";
            var startIndex = source.PartContains("efghijk");

            Assert.Equal(4, startIndex);
        }
    }
}
