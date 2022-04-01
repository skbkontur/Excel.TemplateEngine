using System.IO;

using Benchmark.IhclmeModel;

using BenchmarkDotNet.Attributes;
using BenchmarkDotNet.Jobs;

using SkbKontur.Excel.TemplateEngine;
using SkbKontur.Excel.TemplateEngine.FileGenerating;
using SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

using Vostok.Logging.Abstractions;

namespace Benchmark
{
    [MemoryDiagnoser]
    [SimpleJob(RuntimeMoniker.Net472)]
    [SimpleJob(RuntimeMoniker.Net60)]
    public class ParseVsLazyParse
    {
        [GlobalSetup]
        public void GlobalSetup()
        {
            var templateBytes = File.ReadAllBytes("ihclme_template.xlsx");
            templateDocument = ExcelDocumentFactory.CreateFromTemplate(templateBytes, logger);
            var template = new ExcelTable(templateDocument.GetWorksheet(0));
            templateEngine = new TemplateEngine(template, logger);
        }

        [GlobalCleanup]
        public void GlobalCleanup()
        {
            templateDocument.Dispose();
        }

        [IterationSetup]
        public void IterationSetup()
        {
            targetFileContent = File.ReadAllBytes($"ihclmeWith{ItemsCount}List.xlsx");
        }

        [Benchmark]
        public void Parse()
        {
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(targetFileContent, logger))
            {
                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), logger);
                var tableParser = new TableParser(target, tableNavigator);

                templateEngine.Parse<IhclmeExcelModel>(tableParser);
            }
        }

        [Benchmark]
        public void LazyParse()
        {
            using (var lazyTableReader = new LazyTableReader(targetFileContent))
            {
                templateEngine.LazyParse<IhclmeExcelModel>(lazyTableReader);
            }
        }

        private byte[] targetFileContent;

        private IExcelDocument templateDocument;
        private TemplateEngine templateEngine;

        [Params("500", "5k", "10k")]
        public string ItemsCount;

        private readonly ILog logger = new SilentLog();
    }
}