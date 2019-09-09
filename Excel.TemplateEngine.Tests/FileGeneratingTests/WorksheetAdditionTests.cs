using System.IO;

using NUnit.Framework;

namespace Excel.TemplateEngine.Tests.FileGeneratingTests
{
    public class WorksheetAdditionTests : FileBasedTestBase
    {
        [Test]
        public void WorksheetAdditionTest()
        {
            var document = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), Log.DefaultLogger);
            document.AddWorksheet("Лист2");
            var worksheet = document.GetWorksheet(1);
            Assert.AreNotEqual(null, worksheet);

            var result = document.CloseAndGetDocumentBytes();
            File.WriteAllBytes("output.xlsx", result);

            document.Dispose();
        }
    }
}