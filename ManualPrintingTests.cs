using System;
using System.IO;
using System.Linq;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;
using SKBKontur.Catalogue.ExcelObjectPrinter;
using SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [Ignore("These tests can be used only for manual testing")]
    public class ManualPrintingTests
    {
        [Test]
        public void TestPrintingDropDownFromTheOtherWorksheet()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/printingDropDownFromTheOtherWorksheet.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/empty.xlsx")))
            {
                targetDocument.CopyVbaInfoFrom(templateDocument);

                foreach (var index in Enumerable.Range(1, templateDocument.GetWorksheetCount() - 1))
                {
                    var worksheet = templateDocument.GetWorksheet(index);
                    var name = templateDocument.GetWorksheetName(index);
                    var innterTemplateEngine = new TemplateEngine(new ExcelTable(worksheet));
                    var targetWorksheet = targetDocument.AddWorksheet(name);
                    var innerTableBuilder = new TableBuilder(new ExcelTable(targetWorksheet), new TableNavigator(new CellPosition("A1")));
                    innterTemplateEngine.Render(innerTableBuilder, new { });
                }

                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"));
                var tableBuilder = new TableBuilder(target, tableNavigator, new Style(template.GetCell(new CellPosition("A1"))));
                templateEngine.Render(tableBuilder, new { Type = "Значение 2" });

                var filename = "output.xlsx";
                File.WriteAllBytes(filename, targetDocument.CloseAndGetDocumentBytes());

                var path = "file:///" + Path.GetFullPath(filename).Replace("\\", "/");
                Assert.Fail($"Please manually open file '{path}' and check that dropdown on the first sheet has value 'Значение 2'");
            }
        }

        [Test]
        public void TestPrintingVbaMacros()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/printingVbaMacros.xlsm")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/empty.xlsm")))
            {
                targetDocument.CopyVbaInfoFrom(templateDocument);

                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"));
                var tableBuilder = new TableBuilder(target, tableNavigator, new Style(template.GetCell(new CellPosition("A1"))));
                templateEngine.Render(tableBuilder, new {Type = "123"});

                var filename = "output.xlsm";
                File.WriteAllBytes(filename, targetDocument.CloseAndGetDocumentBytes());

                var path = "file:///" + Path.GetFullPath(filename).Replace("\\", "/");
                Assert.Fail($"Please manually open file '{path}' and check that clicking on the right checkbox leads to changes in both checkbox");
            }
        }

        [Test]
        public void TestColumnsSwitching()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/columnsSwitching.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/empty.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"));
                var tableBuilder = new TableBuilder(target, tableNavigator, new Style(template.GetCell(new CellPosition("A1"))));
                templateEngine.Render(tableBuilder, new { A = "First", B = true, C = "Third", D = new [] {1, 2, 3}, E = "Fifth" });

                var filename = "output.xlsx";
                File.WriteAllBytes(filename, targetDocument.CloseAndGetDocumentBytes());

                var path = "file:///" + Path.GetFullPath(filename).Replace("\\", "/");
                Assert.Fail($"Please manually open file '{path}' and check values:\n" +
                            string.Join("\n", new[]
                                {
                                    "C5: First",
                                    "D5: <empty>",
                                    "E5: Third",
                                    "F5: 1",
                                    "F6: 2",
                                    "F7: 3",
                                    "G5: Fifth",
                                    "Флажок 1: checked",
                                }) + "\n");
            }
        }
    }
}