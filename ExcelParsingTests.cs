using System.Collections.Generic;
using System.IO;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;
using SKBKontur.Catalogue.ExcelObjectPrinter;
using SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class ExcelParsingTests
    {
        // todo (mpivko, 25.12.2017): maybe merge some tests
        // todo (mpivko, 25.12.2017): refactor these test

        [Test]
        public void TestSimpleWithEnumerable()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/simpleWithEnumerable_template.xlsx"));
            var template = new ExcelTable(templateDocument.GetWorksheet(0));
            var templateEngine = new TemplateEngine(template);

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/simpleWithEnumerable_target.xlsx"));

            var target = new ExcelTable(targetDocument.GetWorksheet(0));
            var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
            var tableParser = new TableParser(tableNavigator);
            var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

            Assert.AreEqual("C3", mappingForErrors["Type"]);
            Assert.AreEqual("B13", mappingForErrors["Items[0].Id"]);
            Assert.AreEqual("C13", mappingForErrors["Items[0].Name"]);
            Assert.AreEqual("B14", mappingForErrors["Items[1].Id"]);
            Assert.AreEqual("C14", mappingForErrors["Items[1].Name"]);

            Assert.AreEqual("Основной", model.Type);
            Assert.AreEqual(new[]
                {
                    new Item {Id = "2311129000009", Name = "СЫР ГОЛЛАНДСКИЙ МОЖГА 1КГ"},
                    new Item {Id = "2311131000004", Name = "СЫР РОССИЙСКИЙ МОЖГА 1КГ"},
                }, model.Items);
        }

        [Test]
        public void TestCheckBoxes()
        {
            using(var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/сheckBoxes_template.xlsx")))
            using(var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/сheckBoxes_target.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

                Assert.AreEqual("CheckBoxName1", mappingForErrors["TestFlag1"]);
                Assert.AreEqual("CheckBoxName2", mappingForErrors["TestFlag2"]);
                Assert.AreEqual(false, model.TestFlag1);
                Assert.AreEqual(true, model.TestFlag2);
            }
        }

        [Test]
        public void TestNonexistentField()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/nonexistentField_template.xlsx"));
            var template = new ExcelTable(templateDocument.GetWorksheet(0));
            var templateEngine = new TemplateEngine(template);

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/empty.xlsx"));

            var target = new ExcelTable(targetDocument.GetWorksheet(0));
            var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
            var tableParser = new TableParser(tableNavigator);

            Assert.Throws<InvalidExcelTemplateException>(() => templateEngine.Parse<PriceList>(tableParser));
        }

        [Test]
        public void TestDictDirectAccess()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dictDirectAccess_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dictDirectAccess_target.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

                // todo (mpivko, 25.12.2017): add non-leaf dicts

                Assert.AreEqual("C7", mappingForErrors["StringStringDict[\"testKey\"]"]);
                Assert.AreEqual("E7", mappingForErrors["IntStringDict[42]"]);
                Assert.AreEqual(new Dictionary<string, string> { { "testKey", "testValue" } }, model.StringStringDict);
                Assert.AreEqual(new Dictionary<int, string> { { 42, "testValueInt" } }, model.IntStringDict);
            }
        }

        [Test]
        public void TestDictDirectAccessInCheckBoxes()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dictDirectAccessInCheckBoxes_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dictDirectAccessInCheckBoxes_target.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

                Assert.AreEqual("C7", mappingForErrors["StringStringDict[\"testKey\"]"]);
                Assert.AreEqual("E7", mappingForErrors["IntStringDict[42]"]);
                Assert.AreEqual("TestCheckBox1", mappingForErrors["IntBoolDict[25]"]);
                Assert.AreEqual("TestCheckBox2", mappingForErrors["IntBoolDict[27]"]);
                Assert.AreEqual(new Dictionary<string, string> { { "testKey", "testValue" } }, model.StringStringDict);
                Assert.AreEqual(new Dictionary<int, string> { { 42, "testValueInt" } }, model.IntStringDict);
                Assert.AreEqual(new Dictionary<int, bool> { { 25, false }, { 27, true } }, model.IntBoolDict);
            }
        }

        [Test]
        public void TestDictDirectAccessInDropDown()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dictDirectAccessInDropDown_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dictDirectAccessInDropDown_target.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

                Assert.AreEqual("C7", mappingForErrors["StringStringDict[\"testKey\"]"]);
                Assert.AreEqual("E7", mappingForErrors["IntStringDict[42]"]);
                Assert.AreEqual("TestDropDown1", mappingForErrors["IntStringDict[15]"]);
                Assert.AreEqual(new Dictionary<string, string> { { "testKey", "testValue" } }, model.StringStringDict);
                Assert.AreEqual(new Dictionary<int, string> {{42, "testValueInt"}, {15, "Value2"}}, model.IntStringDict);
            }
        }

        [Test]
        public void TestEmptyEnumerable()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/emptyEnumerable_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/emptyEnumerable_target.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

                Assert.AreEqual("C3", mappingForErrors["Type"]);
                Assert.AreEqual(202, mappingForErrors.Count); // todo (mpivko, 25.12.2017): 

                Assert.AreEqual(0, model.Items.Length);
                Assert.AreEqual("", model.Type);
            }
        }

        [Test]
        public void TestDropDownFromTheOtherWorksheet()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dropDownOtherWorksheet_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/dropDownOtherWorksheet_target.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

                Assert.AreEqual("DropDown1", mappingForErrors["Type"]);

                Assert.AreEqual("ValueC", model.Type);
            }
        }

        [Test]
        public void TestImportAfterCreate()
        {
            byte[] bytes;

            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/importAfterCreate_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/empty.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("A1"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableBuilder = new TableBuilder(tableNavigator);

                templateEngine.Render(tableBuilder, new { Type = "Значение 2", TestFlag1 = false, TestFlag2 = true });
                
                target.InsertCell(new CellPosition("C16"));

                bytes = targetDocument.CloseAndGetDocumentBytes();
            }
            
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/importAfterCreate_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(bytes))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);
                
                Assert.AreEqual("CheckBoxName1", mappingForErrors["TestFlag1"]);
                Assert.AreEqual("CheckBoxName2", mappingForErrors["TestFlag2"]);
                Assert.AreEqual("C3", mappingForErrors["Type"]);
                Assert.AreEqual(false, model.TestFlag1);
                Assert.AreEqual(true, model.TestFlag2);
                Assert.AreEqual("Значение 2", model.Type);
            }
        }

        // todo (mpivko, 25.12.2017): test for xlsm
    }

    #region TestModels

    class Item
    {
        public int Index { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }

        public override string ToString()
        {
            return $"{nameof(Index)}: {Index}, {nameof(Id)}: {Id}, {nameof(Name)}: {Name}";
        }

        protected bool Equals(Item other)
        {
            return Index == other.Index && string.Equals(Id, other.Id) && string.Equals(Name, other.Name);
        }

        public override bool Equals(object obj)
        {
            if (ReferenceEquals(null, obj)) return false;
            if (ReferenceEquals(this, obj)) return true;
            if (obj.GetType() != this.GetType()) return false;
            return Equals((Item)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                var hashCode = Index;
                hashCode = (hashCode * 397) ^ (Id != null ? Id.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Name != null ? Name.GetHashCode() : 0);
                return hashCode;
            }
        }
    }

    class PriceList
    {
        public string Type { get; set; }
        public Item[] Items { get; set; }
        public bool TestFlag1 { get; set; }
        public bool TestFlag2 { get; set; }
        public Dictionary<string, string> StringStringDict { get; set; }
        public Dictionary<int, string> IntStringDict { get; set; }
        public Dictionary<int, bool> IntBoolDict { get; set; }
    }

    #endregion
}