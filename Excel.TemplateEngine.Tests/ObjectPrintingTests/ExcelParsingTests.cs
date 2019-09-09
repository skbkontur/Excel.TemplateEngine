using System.Collections.Generic;
using System.IO;

using Excel.TemplateEngine.FileGenerating;
using Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitivesImplementation;
using Excel.TemplateEngine.ObjectPrinting.Exceptions;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using Excel.TemplateEngine.ObjectPrinting.TableNavigator;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

using NUnit.Framework;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    [TestFixture]
    public class ExcelParsingTests : FileBasedTestBase
    {
        [Test]
        public void TestSimpleWithEnumerable()
        {
            var (model, mappingForErrors) = Parse("simpleWithEnumerable_template.xlsx", "simpleWithEnumerable_target.xlsx");

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
        public void TestEnumerableWithPrimaryKey()
        {
            var (model, mappingForErrors) = Parse("enumerableWithPrimaryKey_template.xlsx", "enumerableWithPrimaryKey_target.xlsx");

            Assert.AreEqual("C3", mappingForErrors["Type"]);
            for (var i = 0; i < 7; i++)
            {
                Assert.AreEqual($"B{i + 13}", mappingForErrors[$"Items[{i}].Id"]);
                Assert.AreEqual($"C{i + 13}", mappingForErrors[$"Items[{i}].Name"]);
                Assert.AreEqual($"D{i + 13}", mappingForErrors[$"Items[{i}].BuyerProductId"]);
                Assert.AreEqual($"E{i + 13}", mappingForErrors[$"Items[{i}].Articul"]);
            }

            Assert.AreEqual("Основной", model.Type);
            Assert.AreEqual(new[]
                {
                    new Item {Id = "2311129000009", Name = "СЫР ГОЛЛАНДСКИЙ МОЖГА 1КГ", BuyerProductId = "000074467", Articul = "123456"},
                    new Item {Id = "2311131000004", Name = "СЫР РОССИЙСКИЙ МОЖГА 1КГ", BuyerProductId = "000074468", Articul = "123457"},
                    new Item {Id = null, Name = "Товар 3"},
                    new Item {Id = "123", Articul = "3123123"},
                    new Item {Id = "111", Articul = "111111"},
                    new Item {Name = "Товар 6"},
                    new Item {Id = "222", Name = "Товар 7", Articul = "123"},
                }, model.Items);
        }

        [Test]
        public void TestCheckBoxes()
        {
            var (model, mappingForErrors) = Parse("сheckBoxes_template.xlsx", "сheckBoxes_target.xlsx");
            Assert.AreEqual("CheckBoxName1", mappingForErrors["TestFlag1"]);
            Assert.AreEqual("CheckBoxName2", mappingForErrors["TestFlag2"]);
            Assert.AreEqual(false, model.TestFlag1);
            Assert.AreEqual(true, model.TestFlag2);
        }

        [Test]
        public void TestNonexistentCheckBoxes()
        {
            var (model, mappingForErrors) = Parse("nonexistent_сheckBoxes_template.xlsx", "nonexistent_сheckBoxes_target.xlsx");
            Assert.AreEqual("CheckBoxName1", mappingForErrors["TestFlag1"]);
            Assert.AreEqual("NonexistentInTarget", mappingForErrors["TestFlag2"]);
            Assert.AreEqual(true, model.TestFlag1);
            Assert.AreEqual(false, model.TestFlag2);
        }

        [Test]
        public void TestNonexistentField()
        {
            Assert.Throws<ObjectPropertyExtractionException>(() => Parse("nonexistentField_template.xlsx", "empty.xlsx"));
        }

        [Test]
        public void TestDictDirectAccess()
        {
            var (model, mappingForErrors) = Parse("dictDirectAccess_template.xlsx", "dictDirectAccess_target.xlsx");
            Assert.AreEqual("C7", mappingForErrors["StringStringDict[\"testKey\"]"]);
            Assert.AreEqual("E7", mappingForErrors["IntStringDict[42]"]);
            Assert.AreEqual("E9", mappingForErrors["InnerPriceList.IntStringDict[25]"]);
            Assert.AreEqual("E10", mappingForErrors["InnerPriceList.PriceListsDict[\"price\"].Type"]);
            Assert.AreEqual(new Dictionary<string, string> {{"testKey", "testValue"}}, model.StringStringDict);
            Assert.AreEqual(new Dictionary<int, string> {{42, "testValueInt"}}, model.IntStringDict);
            Assert.AreEqual(new Dictionary<int, string> {{25, "222222222"}}, model.InnerPriceList.IntStringDict);
            Assert.AreEqual(1, model.InnerPriceList.PriceListsDict.Count);
            Assert.AreEqual("720306001", model.InnerPriceList.PriceListsDict["price"].Type);
        }

        [Test]
        public void TestDictDirectAccessInCheckBoxes()
        {
            var (model, mappingForErrors) = Parse("dictDirectAccessInCheckBoxes_template.xlsx", "dictDirectAccessInCheckBoxes_target.xlsx");
            Assert.AreEqual("C7", mappingForErrors["StringStringDict[\"testKey\"]"]);
            Assert.AreEqual("E7", mappingForErrors["IntStringDict[42]"]);
            Assert.AreEqual("TestCheckBox1", mappingForErrors["IntBoolDict[25]"]);
            Assert.AreEqual("TestCheckBox2", mappingForErrors["IntBoolDict[27]"]);
            Assert.AreEqual(new Dictionary<string, string> {{"testKey", "testValue"}}, model.StringStringDict);
            Assert.AreEqual(new Dictionary<int, string> {{42, "testValueInt"}}, model.IntStringDict);
            Assert.AreEqual(new Dictionary<int, bool> {{25, false}, {27, true}}, model.IntBoolDict);
        }

        [Test]
        public void TestDictDirectAccessInDropDown()
        {
            var (model, mappingForErrors) = Parse("dictDirectAccessInDropDown_template.xlsx", "dictDirectAccessInDropDown_target.xlsx");
            Assert.AreEqual("C7", mappingForErrors["StringStringDict[\"testKey\"]"]);
            Assert.AreEqual("E7", mappingForErrors["IntStringDict[42]"]);
            Assert.AreEqual("TestDropDown1", mappingForErrors["IntStringDict[15]"]);
            Assert.AreEqual(new Dictionary<string, string> {{"testKey", "testValue"}}, model.StringStringDict);
            Assert.AreEqual(new Dictionary<int, string> {{42, "testValueInt"}, {15, "Value2"}}, model.IntStringDict);
        }

        [Test]
        public void TestEmptyEnumerable()
        {
            var (model, mappingForErrors) = Parse("emptyEnumerable_template.xlsx", "emptyEnumerable_target.xlsx");

            Assert.AreEqual("C3", mappingForErrors["Type"]);
            Assert.AreEqual(1, mappingForErrors.Count);

            Assert.AreEqual(0, model.Items.Length);
            Assert.AreEqual("", model.Type);
        }

        [Test]
        public void TestDropDownFromTheOtherWorksheet()
        {
            var (model, mappingForErrors) = Parse("dropDownOtherWorksheet_template.xlsx", "dropDownOtherWorksheet_target.xlsx");
            Assert.AreEqual("DropDown1", mappingForErrors["Type"]);
            Assert.AreEqual("ValueC", model.Type);
        }

        [TestCase("xlsx")]
        [TestCase("xlsm")]
        public void TestImportAfterCreate(string extension)
        {
            byte[] bytes;

            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("importAfterCreate_template.xlsx")), Log.DefaultLogger))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath($"empty.{extension}")), Log.DefaultLogger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, Log.DefaultLogger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), Log.DefaultLogger);
                var tableBuilder = new TableBuilder(target, tableNavigator, new Style(template.GetCell(new CellPosition("A1"))));

                templateEngine.Render(tableBuilder, new {Type = "Значение 2", TestFlag1 = false, TestFlag2 = true});

                target.InsertCell(new CellPosition("C16"));

                bytes = targetDocument.CloseAndGetDocumentBytes();
            }

            var (model, mappingForErrors) = Parse(File.ReadAllBytes(GetFilePath("importAfterCreate_template.xlsx")), bytes);
            Assert.AreEqual("CheckBoxName1", mappingForErrors["TestFlag1"]);
            Assert.AreEqual("CheckBoxName2", mappingForErrors["TestFlag2"]);
            Assert.AreEqual("C3", mappingForErrors["Type"]);
            Assert.AreEqual(false, model.TestFlag1);
            Assert.AreEqual(true, model.TestFlag2);
            Assert.AreEqual("Значение 2", model.Type);
        }

        private (PriceList model, Dictionary<string, string> mappingForErrors) Parse(string templateFileName, string targetFileName)
        {
            return Parse(File.ReadAllBytes(GetFilePath(templateFileName)), File.ReadAllBytes(GetFilePath(targetFileName)));
        }

        private (PriceList model, Dictionary<string, string> mappingForErrors) Parse(byte[] templateBytes, byte[] targetBytes)
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(templateBytes, Log.DefaultLogger))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(targetBytes, Log.DefaultLogger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, Log.DefaultLogger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), Log.DefaultLogger);
                var tableParser = new TableParser(target, tableNavigator);
                return templateEngine.Parse<PriceList>(tableParser);
            }
        }
    }

    #region TestModels

    class Item
    {
        public int Index { get; set; }
        public string Id { get; set; }
        public string Name { get; set; }
        public string BuyerProductId { get; set; }
        public string Articul { get; set; }

        public override string ToString()
        {
            return $"{nameof(Index)}: {Index}, {nameof(Id)}: {Id}, {nameof(Name)}: {Name}, {nameof(BuyerProductId)}: {BuyerProductId}, {nameof(Articul)}: {Articul}";
        }

        protected bool Equals(Item other)
        {
            return Index == other.Index && string.Equals(Id, other.Id) && string.Equals(Name, other.Name) && string.Equals(BuyerProductId, other.BuyerProductId) && string.Equals(Articul, other.Articul);
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
                hashCode = (hashCode * 397) ^ (BuyerProductId != null ? BuyerProductId.GetHashCode() : 0);
                hashCode = (hashCode * 397) ^ (Articul != null ? Articul.GetHashCode() : 0);
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
        public PriceList InnerPriceList { get; set; }
        public Dictionary<string, PriceList> PriceListsDict { get; set; }
    }

    #endregion
}