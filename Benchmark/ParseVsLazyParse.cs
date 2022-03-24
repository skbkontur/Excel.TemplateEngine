using System.IO;

using Benchmark.IhclmeModel;

using BenchmarkDotNet.Attributes;

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
    public class ParseVsLazyParse
    {
        private IExcelDocument GetTemplate()
        {
            var templateBytes = File.ReadAllBytes("ihclme_template.xlsx");
            return ExcelDocumentFactory.CreateFromTemplate(templateBytes, logger);
        }

        [Benchmark]
        public void Parse5k()
        {
            var targetBytes = File.ReadAllBytes("ihclmeWith5kList.xlsx");

            using (var templateDocument = GetTemplate())
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(targetBytes, logger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), logger);
                var tableParser = new TableParser(target, tableNavigator);
                templateEngine.Parse<IhclmeExcelModel>(tableParser);
            }
        }

        [Benchmark]
        public void Parse10k()
        {
            var targetBytes = File.ReadAllBytes("ihclmeWith10kList.xlsx");

            using (var templateDocument = GetTemplate())
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(targetBytes, logger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), logger);
                var tableParser = new TableParser(target, tableNavigator);
                templateEngine.Parse<IhclmeExcelModel>(tableParser);
            }
        }

        [Benchmark]
        public void Parse500()
        {
            var targetBytes = File.ReadAllBytes("ihclmeWith500List.xlsx");

            using (var templateDocument = GetTemplate())
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(targetBytes, logger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), logger);
                var tableParser = new TableParser(target, tableNavigator);
                templateEngine.Parse<IhclmeExcelModel>(tableParser);
            }
        }

        [Benchmark]
        public void LazyParse500()
        {
            using (var templateDocument = GetTemplate())
            {
                var targetBytes = File.ReadAllBytes("ihclmeWith500List.xlsx");

                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var lazyTableReader = new LazyTableReader(targetBytes);
                templateEngine.LazyParse<IhclmeExcelModel>(lazyTableReader);
            }
        }

        [Benchmark]
        public void LazyParse5k()
        {
            using (var templateDocument = GetTemplate())
            {
                var targetBytes = File.ReadAllBytes("ihclmeWith5kList.xlsx");

                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var lazyTableReader = new LazyTableReader(targetBytes);
                templateEngine.LazyParse<IhclmeExcelModel>(lazyTableReader);
            }
        }

        [Benchmark]
        public void LazyParse10k()
        {
            using (var templateDocument = GetTemplate())
            {
                var targetBytes = File.ReadAllBytes("ihclmeWith10kList.xlsx");

                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var lazyTableReader = new LazyTableReader(targetBytes);
                templateEngine.LazyParse<IhclmeExcelModel>(lazyTableReader);
            }
        }

        private readonly ILog logger = new SilentLog();
    }
}