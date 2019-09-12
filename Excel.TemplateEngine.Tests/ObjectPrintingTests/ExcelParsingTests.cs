using System;
using System.Collections.Generic;
using System.IO;

using Excel.TemplateEngine.Exceptions;
using Excel.TemplateEngine.FileGenerating;
using Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using Excel.TemplateEngine.ObjectPrinting.TableNavigator;
using Excel.TemplateEngine.ObjectPrinting.TableParser;

using FluentAssertions;

using NUnit.Framework;

using Vostok.Logging.Console;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    [TestFixture]
    public class ExcelParsingTests : FileBasedTestBase
    {
        [Test]
        public void TestSimpleWithEnumerable()
        {
            var (model, mappingForErrors) = Parse("simpleWithEnumerable_template.xlsx", "simpleWithEnumerable_target.xlsx");

            mappingForErrors["Type"].Should().Be("C3");
            mappingForErrors["Items[0].Id"].Should().Be("B13");
            mappingForErrors["Items[0].Name"].Should().Be("C13");
            mappingForErrors["Items[1].Id"].Should().Be("B14");
            mappingForErrors["Items[1].Name"].Should().Be("C14");

            model.Type.Should().Be("Основной");
            model.Items.Should().BeEquivalentTo(new[]
                {
                    new Item {Id = "2311129000009", Name = "СЫР ГОЛЛАНДСКИЙ МОЖГА 1КГ"},
                    new Item {Id = "2311131000004", Name = "СЫР РОССИЙСКИЙ МОЖГА 1КГ"},
                });
        }

        [Test]
        public void TestEnumerableWithPrimaryKey()
        {
            var (model, mappingForErrors) = Parse("enumerableWithPrimaryKey_template.xlsx", "enumerableWithPrimaryKey_target.xlsx");

            mappingForErrors["Type"].Should().Be("C3");
            for (var i = 0; i < 7; i++)
            {
                mappingForErrors[$"Items[{i}].Id"].Should().Be($"B{i + 13}");
                mappingForErrors[$"Items[{i}].Name"].Should().Be($"C{i + 13}");
                mappingForErrors[$"Items[{i}].BuyerProductId"].Should().Be($"D{i + 13}");
                mappingForErrors[$"Items[{i}].Articul"].Should().Be($"E{i + 13}");
            }

            model.Type.Should().Be("Основной");
            model.Items.Should().BeEquivalentTo(
                new Item {Id = "2311129000009", Name = "СЫР ГОЛЛАНДСКИЙ МОЖГА 1КГ", BuyerProductId = "000074467", Articul = "123456"},
                new Item {Id = "2311131000004", Name = "СЫР РОССИЙСКИЙ МОЖГА 1КГ", BuyerProductId = "000074468", Articul = "123457"},
                new Item {Id = null, Name = "Товар 3"},
                new Item {Id = "123", Articul = "3123123"},
                new Item {Id = "111", Articul = "111111"},
                new Item {Name = "Товар 6"},
                new Item {Id = "222", Name = "Товар 7", Articul = "123"}
            );
        }

        [Test]
        public void TestCheckBoxes()
        {
            var (model, mappingForErrors) = Parse("сheckBoxes_template.xlsx", "сheckBoxes_target.xlsx");
            mappingForErrors["TestFlag1"].Should().Be("CheckBoxName1");
            mappingForErrors["TestFlag2"].Should().Be("CheckBoxName2");
            model.TestFlag1.Should().BeFalse();
            model.TestFlag2.Should().BeTrue();
        }

        [Test]
        public void TestNonexistentCheckBoxes()
        {
            var (model, mappingForErrors) = Parse("nonexistent_сheckBoxes_template.xlsx", "nonexistent_сheckBoxes_target.xlsx");
            mappingForErrors["TestFlag1"].Should().Be("CheckBoxName1");
            mappingForErrors["TestFlag2"].Should().Be("NonexistentInTarget");
            model.TestFlag1.Should().BeTrue();
            model.TestFlag2.Should().BeFalse();
        }

        [Test]
        public void TestNonexistentField()
        {
            Action parsing = () => Parse("nonexistentField_template.xlsx", "empty.xlsx");
            parsing.Should().Throw<ObjectPropertyExtractionException>();
        }

        [Test]
        public void TestDictDirectAccess()
        {
            var (model, mappingForErrors) = Parse("dictDirectAccess_template.xlsx", "dictDirectAccess_target.xlsx");
            mappingForErrors["StringStringDict[\"testKey\"]"].Should().Be("C7");
            mappingForErrors["IntStringDict[42]"].Should().Be("E7");
            mappingForErrors["InnerPriceList.IntStringDict[25]"].Should().Be("E9");
            mappingForErrors["InnerPriceList.PriceListsDict[\"price\"].Type"].Should().Be("E10");
            model.StringStringDict.Should().BeEquivalentTo(new Dictionary<string, string> {{"testKey", "testValue"}});
            model.IntStringDict.Should().BeEquivalentTo(new Dictionary<int, string> {{42, "testValueInt"}});
            model.InnerPriceList.IntStringDict.Should().BeEquivalentTo(new Dictionary<int, string> {{25, "222222222"}});
            model.InnerPriceList.PriceListsDict.Count.Should().Be(1);
            model.InnerPriceList.PriceListsDict["price"].Type.Should().Be("720306001");
        }

        [Test]
        public void TestDictDirectAccessInCheckBoxes()
        {
            var (model, mappingForErrors) = Parse("dictDirectAccessInCheckBoxes_template.xlsx", "dictDirectAccessInCheckBoxes_target.xlsx");
            mappingForErrors["StringStringDict[\"testKey\"]"].Should().Be("C7");
            mappingForErrors["IntStringDict[42]"].Should().Be("E7");
            mappingForErrors["IntBoolDict[25]"].Should().Be("TestCheckBox1");
            mappingForErrors["IntBoolDict[27]"].Should().Be("TestCheckBox2");
            model.StringStringDict.Should().BeEquivalentTo(new Dictionary<string, string> {{"testKey", "testValue"}});
            model.IntStringDict.Should().BeEquivalentTo(new Dictionary<int, string> {{42, "testValueInt"}});
            model.IntBoolDict.Should().BeEquivalentTo(new Dictionary<int, bool> {{25, false}, {27, true}});
        }

        [Test]
        public void TestDictDirectAccessInDropDown()
        {
            var (model, mappingForErrors) = Parse("dictDirectAccessInDropDown_template.xlsx", "dictDirectAccessInDropDown_target.xlsx");
            mappingForErrors["StringStringDict[\"testKey\"]"].Should().Be("C7");
            mappingForErrors["IntStringDict[42]"].Should().Be("E7");
            mappingForErrors["IntStringDict[15]"].Should().Be("TestDropDown1");
            model.StringStringDict.Should().BeEquivalentTo(new Dictionary<string, string> {{"testKey", "testValue"}});
            model.IntStringDict.Should().BeEquivalentTo(new Dictionary<int, string> {{42, "testValueInt"}, {15, "Value2"}});
        }

        [Test]
        public void TestEmptyEnumerable()
        {
            var (model, mappingForErrors) = Parse("emptyEnumerable_template.xlsx", "emptyEnumerable_target.xlsx");

            mappingForErrors["Type"].Should().Be("C3");
            mappingForErrors.Count.Should().Be(1);

            model.Items.Length.Should().Be(0);
            model.Type.Should().BeEmpty();
        }

        [Test]
        public void TestDropDownFromTheOtherWorksheet()
        {
            var (model, mappingForErrors) = Parse("dropDownOtherWorksheet_template.xlsx", "dropDownOtherWorksheet_target.xlsx");
            mappingForErrors["Type"].Should().Be("DropDown1");
            model.Type.Should().Be("ValueC");
        }

        [TestCase("xlsx")]
        [TestCase("xlsm")]
        public void TestImportAfterCreate(string extension)
        {
            byte[] bytes;

            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("importAfterCreate_template.xlsx")), logger))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath($"empty.{extension}")), logger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), logger);
                var tableBuilder = new TableBuilder(target, tableNavigator, new Style(template.GetCell(new CellPosition("A1"))));

                templateEngine.Render(tableBuilder, new {Type = "Значение 2", TestFlag1 = false, TestFlag2 = true});

                target.InsertCell(new CellPosition("C16"));

                bytes = targetDocument.CloseAndGetDocumentBytes();
            }

            var (model, mappingForErrors) = Parse(File.ReadAllBytes(GetFilePath("importAfterCreate_template.xlsx")), bytes);
            mappingForErrors["TestFlag1"].Should().Be("CheckBoxName1");
            mappingForErrors["TestFlag2"].Should().Be("CheckBoxName2");
            mappingForErrors["Type"].Should().Be("C3");
            model.TestFlag1.Should().BeFalse();
            model.TestFlag2.Should().BeTrue();
            model.Type.Should().Be("Значение 2");
        }

        private (PriceList model, Dictionary<string, string> mappingForErrors) Parse(string templateFileName, string targetFileName)
        {
            return Parse(File.ReadAllBytes(GetFilePath(templateFileName)), File.ReadAllBytes(GetFilePath(targetFileName)));
        }

        private (PriceList model, Dictionary<string, string> mappingForErrors) Parse(byte[] templateBytes, byte[] targetBytes)
        {
            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(templateBytes, logger))
            using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(targetBytes, logger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                var target = new ExcelTable(targetDocument.GetWorksheet(0));
                var tableNavigator = new TableNavigator(new CellPosition("A1"), logger);
                var tableParser = new TableParser(target, tableNavigator);
                return templateEngine.Parse<PriceList>(tableParser);
            }
        }

        private readonly ConsoleLog logger = new ConsoleLog();
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