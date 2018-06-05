using FluentAssertions;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives;

namespace SKBKontur.Catalogue.Core.Tests.ExcelFileGeneratorTests
{
    [TestFixture]
    public class ExcelRowTests
    {
        [Test]
        public void CompareCellRefferencesTest()
        {
            var A1 = "A1";
            var B1 = "B1";
            var AA1 = "AA1";
            var AB1 = "AB1";

            Assert.IsTrue(ExcelRow.CompareCellRefferences(B1, A1), "B1 bigger than A1 ");
            Assert.IsTrue(ExcelRow.CompareCellRefferences(AA1, B1), "AA1 bigger than B1 ");
            Assert.IsTrue(ExcelRow.CompareCellRefferences(AB1, AA1), "AB1 bigger than AA1 ");
            Assert.IsTrue(ExcelRow.CompareCellRefferences(AB1, B1), "AB1 bigger than B1 ");
        }
    }
}