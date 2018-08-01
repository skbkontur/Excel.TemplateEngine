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
            var document = ExcelDocumentFactory.CreateFromTemplate(GetFileBytes());
            document.AddWorksheet("Лист2");
            var worksheet = document.GetWorksheet(1);
            Assert.AreNotEqual(null, worksheet);

            var result = document.CloseAndGetDocumentBytes();
            File.WriteAllBytes("output.xlsx", result);

            document.Dispose();
        }

        private static byte[] GetFileBytes()
        {
            const string fileName = @"ExcelFileGeneratorTests\Files\empty.xlsx";
            return File.ReadAllBytes(TestContext.CurrentContext.TestDirectory + "\\" + fileName);
        }
    }
}