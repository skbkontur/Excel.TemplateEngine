using System.IO;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.FileGenerating;

using Vostok.Logging.Console;

namespace SkbKontur.Excel.TemplateEngine.Tests.FileGeneratingTests
{
    [TestFixture]
    public class TemplateTests : FileBasedTestBase
    {
        [TestCase("template_2.xlsx", TestName = "TemplateWithoutTheme")]
        [TestCase("MyOffice_document.xlsx", TestName = "MyOfficeDocument")]
        public void CreateFromTemplateTest(string filename)
        {
            var content = File.ReadAllBytes(GetFilePath(filename));
            using (var template = ExcelDocumentFactory.TryCreateFromTemplate(content, new ConsoleLog()))
            {
                Assert.IsNotNull(template);
            }
        }
    }
}