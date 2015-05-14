using System.IO;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;

namespace SKBKontur.Catalogue.Core.Tests.ExcelFileGeneratorTests
{
    [TestFixture]
    public class WorksheetAdditionTests
    {
        [Test]
        public void WorksheetAdditionTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(emptyFileName));
            document.AddWorksheet("Лист2");
            var worksheet = document.GetWorksheet(1);
            Assert.AreNotEqual(null, worksheet);

            var result = document.CloseAndGetDocumentBytes();
            File.WriteAllBytes("output.xlsx", result);

            document.Dispose();
        }

        private const string emptyFileName = @"ExcelFileGeneratorTests\Files\empty.xlsx";
    }
}