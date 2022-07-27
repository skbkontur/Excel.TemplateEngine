using System.IO;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.FileGenerating;

using Vostok.Logging.Console;

namespace SkbKontur.Excel.TemplateEngine.Tests.FileGeneratingTests
{
    [TestFixture]
    public class TemplateTests : FileBasedTestBase
    {
        [Test]
        public void TemplateWithoutThemeTest()
        {
            var content = File.ReadAllBytes(GetFilePath("template_2.xlsx"));
            using (var template = ExcelDocumentFactory.TryCreateFromTemplate(content, new ConsoleLog()))
            {
                Assert.IsNotNull(template);
            }
        }
    }
}