using System.Collections.Generic;
using System.IO;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelFileGenerator;
using SKBKontur.Catalogue.ExcelObjectPrinter;
using SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    class Item
    {
        public string Id { get; set; }
        public string Name { get; set; }

        protected bool Equals(Item other)
        {
            return string.Equals(Id, other.Id) && string.Equals(Name, other.Name);
        }

        public override bool Equals(object obj)
        {
            if(ReferenceEquals(null, obj)) return false;
            if(ReferenceEquals(this, obj)) return true;
            if(obj.GetType() != this.GetType()) return false;
            return Equals((Item)obj);
        }

        public override int GetHashCode()
        {
            unchecked
            {
                return ((Id != null ? Id.GetHashCode() : 0) * 397) ^ (Name != null ? Name.GetHashCode() : 0);
            }
        }

        public override string ToString()
        {
            return $"{nameof(Id)}: {Id}, {nameof(Name)}: {Name}";
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

    class RegionsHolder
    {
        public List<string> Regions { get; set; }
    }

    [TestFixture]
    public class Test
    {
        [Test]
        public void A()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_A_template.xlsx"));
            var template = new ExcelTable(templateDocument.GetWorksheet(0));
            var templateEngine = new TemplateEngine(template);

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_A_target.xlsx"));

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
        public void B()
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_B_template.xlsx"));
            var template = new ExcelTable(templateDocument.GetWorksheet(0));
            var templateEngine = new TemplateEngine(template);

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_B_target.xlsx"));

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
        public void С()
        {
            return;
            using(var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/defaultPriceList.xlsx")))
            using(var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/empty.xlsx")))
            {
                targetDocument.AddFormControlInfos(0, templateDocument.GetFormControlInfos(0));

                var bytes = targetDocument.CloseAndGetDocumentBytes();
                File.WriteAllBytes("test.xlsx", bytes);
            }
        }

        [Test]
        public void D()
        {
            using(var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_D_template.xlsx")))
            using(var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_D_target.xlsx")))
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
        public void E()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_E_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_E_target.xlsx")))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(target, new CellPosition("B2"), new Styler(template.GetCell(new CellPosition("A1"))));
                var tableParser = new TableParser(tableNavigator);
                var (model, mappingForErrors) = templateEngine.Parse<PriceList>(tableParser);

                Assert.AreEqual("C7", mappingForErrors["StringStringDict[\"testKey\"]"]);
                Assert.AreEqual("E7", mappingForErrors["IntStringDict[42]"]);
                //Assert.AreEqual("Items", mappingForErrors["Items[4].Id"]);
                //Assert.AreEqual("Items", mappingForErrors["Items[4].Id"]);
                Assert.AreEqual(new Dictionary<string, string> { { "testKey", "testValue" } }, model.StringStringDict);
                Assert.AreEqual(new Dictionary<int, string> { { 42, "testValueInt" } }, model.IntStringDict);
                //Assert.AreEqual(new[] { new Item { Id = "2311129000009", Name = "СЫР ГОЛЛАНДСКИЙ МОЖГА 1КГ" } }, model.Items);
            }
        }

        [Test]
        public void F()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_F_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_F_target.xlsx")))
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
                //Assert.AreEqual("Items", mappingForErrors["Items[4].Id"]);
                //Assert.AreEqual("Items", mappingForErrors["Items[4].Id"]);
                Assert.AreEqual(new Dictionary<string, string> { { "testKey", "testValue" } }, model.StringStringDict);
                Assert.AreEqual(new Dictionary<int, string> { { 42, "testValueInt" } }, model.IntStringDict);
                Assert.AreEqual(new Dictionary<int, bool> { { 25, false }, { 27, true } }, model.IntBoolDict);
                //Assert.AreEqual(new[] { new Item { Id = "2311129000009", Name = "СЫР ГОЛЛАНДСКИЙ МОЖГА 1КГ" } }, model.Items);
            }
        }

        [Test]
        public void G()
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_G_template.xlsx")))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes("ExcelObjectPrinterTests/Files/test_G_target.xlsx")))
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
    }

    [TestFixture]
    public class TemplateDescriptionHelperTests
    {
        [Test]
        public void TemplateDescriptionCorrectnessTest()
        {
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:TemplateName:A1:B2"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:TemplateName:ABCD122:BER22"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:TemplateName:ABCDEFGHIJKLMNOPQRSTUVWXYZ1:B1234567890"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("TemplateTemplateName:A1:B2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription(":TemplateName::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template::A2:BCD33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:B:33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:33:33"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:QWE:EWQ"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template::A1:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:A0:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:A02:A2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription("Template:Name:A02:a2"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectTemplateDescription(""));
        }

        [Test]
        public void ValueDescriptionCorrectnessTest()
        {
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::Property"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::P"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::A.B"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::aa.bb"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::Property[]"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::A[].B[].Cc.ddd[]"));
            Assert.IsTrue(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::B52"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::9das"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::Asdf["));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::asdf]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::A[]B[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::sdf.[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::128[]"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value::asd,vcd"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Value:Template:"));
            Assert.IsFalse(TemplateDescriptionHelper.Instance.IsCorrectValueDescription("Val::A"));
        }

        [Test]
        public void TemplateCoordinatesTest()
        {
            IRectangle range;
            Assert.IsTrue(TemplateDescriptionHelper.Instance.TryExtractCoordinates("Template:qwe:QWE123:ASD987", out range));
            Assert.AreEqual(26 * 26 + 19 * 26 + 4, range.UpperLeft.ColumnIndex);
            Assert.AreEqual(123, range.UpperLeft.RowIndex);
            Assert.AreEqual(17 * 26 * 26 + 23 * 26 + 5, range.LowerRight.ColumnIndex);
            Assert.AreEqual(987, range.LowerRight.RowIndex);
        }
    }
}