using System;
using System.Linq;

using FluentAssertions;

using NUnit.Framework;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;
using SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests.FakeDocumentPrimitivesImplementation;

using Vostok.Logging.Console;

namespace SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    [TestFixture]
    public class ObjectFakePrintingTests
    {
        [Test]
        public void StringValuePrintingTest()
        {
            var template = new FakeTable(1, 1);
            var cell = template.InsertCell(new CellPosition(1, 1));
            cell.StringValue = "Template:RootTemplate:A1:A1";

            var templateEngine = new TemplateEngine(template, logger);

            var target = new FakeTable(10, 10);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger), new Style(new FakeCell(new CellPosition("A1"))
                {
                    StyleId = "(1,1)"
                }));

            templateEngine.Render(tableBuilder, "TestStringValue");

            target.GetCell(new CellPosition("B2")).StringValue.Should().Be("TestStringValue");

            ((FakeCell)target.GetCell(new CellPosition("B2"))).StyleId.Should().Be("(1,1)");
        }

        [Test]
        public void SimpleTestWithArray()
        {
            var model = new[] {"a", "b", "c", "d", "e"};

            var template = new FakeTable(1, 1);
            var cell = template.InsertCell(new CellPosition(1, 1));
            cell.StringValue = "Template:RootTemplate:A1:A1";

            var templateEngine = new TemplateEngine(template, logger);

            var target = new FakeTable(10, 10);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger), new Style(new FakeCell(new CellPosition("A1"))
                {
                    StyleId = "(1,1)"
                }));

            templateEngine.Render(tableBuilder, model);

            target.GetCell(new CellPosition("B2")).StringValue.Should().Be("a");
            target.GetCell(new CellPosition("B3")).StringValue.Should().Be("b");
            target.GetCell(new CellPosition("B4")).StringValue.Should().Be("c");
            target.GetCell(new CellPosition("B5")).StringValue.Should().Be("d");
            target.GetCell(new CellPosition("B6")).StringValue.Should().Be("e");

            DebugPrinting(target, new CellPosition(1, 1), new CellPosition(10, 10));
        }

        [Test]
        public void SimpleObjectPrintingTest()
        {
            var model = new Document
                {
                    Buyer = new Organization
                        {
                            Address = "BuyerAddress",
                            Name = "BuyerName"
                        },
                    Supplier = new Organization
                        {
                            Address = "SupplierAddress",
                            Name = "SupplierName"
                        },
                    TypeName = "ORDERS"
                };

            var stringTemplate = new[]
                {
                    new[] {"", "", "", "", ""},
                    new[] {"", "", "", "", ""},
                    new[] {"", "Template:RootTemplate:B4:D5", "", "", ""},
                    new[] {"", "Покупатель:", "Поставщик:", "", ""},
                    new[] {"", "Value:Organization:Buyer", "Value:Organization:Supplier", "Value::TypeName", ""},
                    new[] {"", "", "", "", ""},
                    new[] {"Template:Organization:A8:A9", "", "", "", ""},
                    new[] {"Value::Name", "", "", "", ""},
                    new[] {"Value::Address", "", "", "", ""}
                };
            var template = FakeTable.GenerateFromStringArray(stringTemplate);

            var target = new FakeTable(100, 100);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("A1"), logger));
            var templateEngine = new TemplateEngine(template, logger);
            templateEngine.Render(tableBuilder, model);

            target.GetCell(new CellPosition("A1")).StringValue.Should().Be("Покупатель:");
            target.GetCell(new CellPosition("B1")).StringValue.Should().Be("Поставщик:");
            target.GetCell(new CellPosition("A2")).StringValue.Should().Be("BuyerName");
            target.GetCell(new CellPosition("B2")).StringValue.Should().Be("SupplierName");
            target.GetCell(new CellPosition("C2")).StringValue.Should().Be("ORDERS");
            target.GetCell(new CellPosition("A3")).StringValue.Should().Be("BuyerAddress");
            target.GetCell(new CellPosition("B3")).StringValue.Should().Be("SupplierAddress");

            ((FakeCell)target.GetCell(new CellPosition("A1"))).StyleId.Should().Be("B4");
            ((FakeCell)target.GetCell(new CellPosition("B1"))).StyleId.Should().Be("C4");
            ((FakeCell)target.GetCell(new CellPosition("A2"))).StyleId.Should().Be("A8");
            ((FakeCell)target.GetCell(new CellPosition("B2"))).StyleId.Should().Be("A8");
            ((FakeCell)target.GetCell(new CellPosition("C2"))).StyleId.Should().Be("D5");
            ((FakeCell)target.GetCell(new CellPosition("A3"))).StyleId.Should().Be("A9");
            ((FakeCell)target.GetCell(new CellPosition("B3"))).StyleId.Should().Be("A9");

            DebugPrinting(target, new CellPosition(1, 1), new CellPosition(10, 10));
        }

        [Test]
        public void WithTemplateDecompositionTest()
        {
            var model = new Document
                {
                    Buyer = new Organization
                        {
                            Address = "BuyerAddress",
                            Name = "BuyerName"
                        },
                    Supplier = new Organization
                        {
                            Address = "SupplierAddress",
                            Name = "SupplierName",
                            Inn = "90238192038",
                            Kpp = "0832309812"
                        },
                    Payer = new Organization
                        {
                            Address = "PayerAddress",
                            Name = "PayerName"
                        },
                    DeliveryParty = new Organization
                        {
                            Address = "DeliveryPartyAddress",
                            Name = "DeliveryPartyName"
                        },
                    Vehicle = new VehicleInfo
                        {
                            NameOfCarrier = "Евкакий",
                            TransportMode = "Agressive",
                            VehicleBrand = "Tauria",
                            VehicleNumber = "A777AB"
                        },
                    TypeName = "ORDERS"
                };

            var stringTemplate = new[]
                {
                    new[] {"Template:RootTemplate:A2:D5", "", "", "", "", ""},
                    new[] {"Покупатель:", "", "Поставщик:", "", "", ""},
                    new[] {"Value:Organization:Buyer", "Value::TypeName", "Value:DetailedOrganization:Supplier", "", "", ""},
                    new[] {"Плательщик:", "Средство доставки:", "", "Доставка:", "", ""},
                    new[] {"Value:Organization:Payer", "Value:VehicleInfo:Vehicle", "Value:Organization:DeliveryParty", "", "", ""},
                    new[] {"", "", "", "", "", ""},
                    new[] {"Template:Organization:A8:A9", "", "", "", "", ""},
                    new[] {"Value::Name", "", "", "", "", ""},
                    new[] {"Value::Address", "", "", "", "", ""},
                    new[] {"", "", "", "", "", ""},
                    new[] {"Template:VehicleInfo:A12:B13", "", "", "", "", ""},
                    new[] {"Value::NameOfCarrier", "Value::VehicleBrand", "", "", "", ""},
                    new[] {"Value::TransportMode", "Value::VehicleNumber", "", "", "", ""},
                    new[] {"", "", "", "", "", ""},
                    new[] {"Template:DetailedOrganization:A16:B17", "", "", "", "", ""},
                    new[] {"Value::Name", "Value::Inn", "", "", "", ""},
                    new[] {"Value::Address", "Value::Kpp", "", "", "", ""}
                };

            var template = FakeTable.GenerateFromStringArray(stringTemplate);

            var target = new FakeTable(100, 100);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("A1"), logger));
            var templateEngine = new TemplateEngine(template, logger);
            templateEngine.Render(tableBuilder, model);

            target.GetCell(new CellPosition("A1")).StringValue.Should().Be("Покупатель:");
            target.GetCell(new CellPosition("C1")).StringValue.Should().Be("Поставщик:");
            target.GetCell(new CellPosition("A2")).StringValue.Should().Be("BuyerName");
            target.GetCell(new CellPosition("A3")).StringValue.Should().Be("BuyerAddress");
            target.GetCell(new CellPosition("B2")).StringValue.Should().Be("ORDERS");
            target.GetCell(new CellPosition("C2")).StringValue.Should().Be("SupplierName");
            target.GetCell(new CellPosition("D2")).StringValue.Should().Be("90238192038");
            target.GetCell(new CellPosition("C3")).StringValue.Should().Be("SupplierAddress");
            target.GetCell(new CellPosition("D3")).StringValue.Should().Be("0832309812");
            target.GetCell(new CellPosition("A4")).StringValue.Should().Be("Плательщик:");
            target.GetCell(new CellPosition("B4")).StringValue.Should().Be("Средство доставки:");
            target.GetCell(new CellPosition("D4")).StringValue.Should().Be("Доставка:");
            target.GetCell(new CellPosition("A5")).StringValue.Should().Be("PayerName");
            target.GetCell(new CellPosition("A6")).StringValue.Should().Be("PayerAddress");
            target.GetCell(new CellPosition("B5")).StringValue.Should().Be("Евкакий");
            target.GetCell(new CellPosition("C5")).StringValue.Should().Be("Tauria");
            target.GetCell(new CellPosition("B6")).StringValue.Should().Be("Agressive");
            target.GetCell(new CellPosition("C6")).StringValue.Should().Be("A777AB");
            target.GetCell(new CellPosition("D5")).StringValue.Should().Be("DeliveryPartyName");
            target.GetCell(new CellPosition("D6")).StringValue.Should().Be("DeliveryPartyAddress");

            DebugPrinting(target, new CellPosition(1, 1), new CellPosition(20, 20));
        }

        [Test]
        public void ObjectWithArrayPrintingTest()
        {
            var model = new DocumentWithArray
                {
                    Array = new[]
                        {
                            new Organization
                                {
                                    Address = "Address1",
                                    Name = "Name1"
                                },
                            new Organization
                                {
                                    Address = "Address2",
                                    Name = "Name2"
                                },
                            new Organization
                                {
                                    Address = "Address3",
                                    Name = "Name3"
                                },
                            new Organization
                                {
                                    Address = "Address4",
                                    Name = "Name4"
                                }
                        },
                    NonArray = "StringValue"
                };
            var stringTemplate = new[]
                {
                    new[] {"", "Template:RootTemplate:A2:D4", "", ""},
                    new[] {"", "Адреса:", "Имена:", ""},
                    new[] {"", "Value::Array[].Address", "Value::Array[].Name", ""},
                    new[] {"Value::NonArray", "", "", ""}
                };

            var template = FakeTable.GenerateFromStringArray(stringTemplate);

            var target = new FakeTable(100, 100);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger));
            var templateEngine = new TemplateEngine(template, logger);
            templateEngine.Render(tableBuilder, model);

            target.GetCell(new CellPosition("C2")).StringValue.Should().Be("Адреса:");
            target.GetCell(new CellPosition("D2")).StringValue.Should().Be("Имена:");
            target.GetCell(new CellPosition("C3")).StringValue.Should().Be("Address1");
            target.GetCell(new CellPosition("C4")).StringValue.Should().Be("Address2");
            target.GetCell(new CellPosition("C5")).StringValue.Should().Be("Address3");
            target.GetCell(new CellPosition("C6")).StringValue.Should().Be("Address4");
            target.GetCell(new CellPosition("D3")).StringValue.Should().Be("Name1");
            target.GetCell(new CellPosition("D4")).StringValue.Should().Be("Name2");
            target.GetCell(new CellPosition("D5")).StringValue.Should().Be("Name3");
            target.GetCell(new CellPosition("D6")).StringValue.Should().Be("Name4");
            target.GetCell(new CellPosition("B7")).StringValue.Should().Be("StringValue");

            DebugPrinting(target, new CellPosition(1, 1), new CellPosition(20, 20));
        }

        [Test]
        public void WithNullFiledsTest()
        {
            var model = new Organization
                {
                    Address = null,
                    Name = "Name",
                };
            var stringTemplate = new[]
                {
                    new[] {"Template:RootTemplate:A2:B2", ""},
                    new[] {"Value::Address", "Value::Name"}
                };

            var template = FakeTable.GenerateFromStringArray(stringTemplate);

            var target = new FakeTable(100, 100);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger));
            var templateEngine = new TemplateEngine(template, logger);
            templateEngine.Render(tableBuilder, model);

            DebugPrinting(target, new CellPosition(1, 1), new CellPosition(20, 20));
        }

        [Test]
        public void WithNullArrayTest()
        {
            var model = new DocumentWithArray
                {
                    Array = null,
                    NonArray = "StringValue"
                };
            var stringTemplate = new[]
                {
                    new[] {"", "Template:RootTemplate:A2:D4", "", ""},
                    new[] {"", "Адреса:", "Имена:", ""},
                    new[] {"", "Value::Array[].Address", "Value::Array[].Name", ""},
                    new[] {"Value::NonArray", "", "", ""}
                };

            var template = FakeTable.GenerateFromStringArray(stringTemplate);

            var target = new FakeTable(100, 100);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger));
            var templateEngine = new TemplateEngine(template, logger);
            templateEngine.Render(tableBuilder, model);

            target.GetCell(new CellPosition("C2")).StringValue.Should().Be("Адреса:");
            target.GetCell(new CellPosition("D2")).StringValue.Should().Be("Имена:");
            target.GetCell(new CellPosition("C3")).StringValue.Should().BeEmpty();
            target.GetCell(new CellPosition("D3")).StringValue.Should().BeEmpty();
            target.GetCell(new CellPosition("B4")).StringValue.Should().Be("StringValue");

            DebugPrinting(target, new CellPosition(1, 1), new CellPosition(20, 20));
        }

        [Test]
        public void ObjectWithArrayAndNullFieldsPrintingTest()
        {
            var model = new DocumentWithArray
                {
                    Array = new[]
                        {
                            new Organization
                                {
                                    Address = "Address1",
                                    Name = "Name1"
                                },
                            new Organization
                                {
                                    Address = "Address2",
                                    Name = null
                                },
                            new Organization
                                {
                                    Address = null,
                                    Name = "Name3"
                                },
                            null,
                            new Organization
                                {
                                    Address = "Address4",
                                    Name = "Name4"
                                }
                        },
                    NonArray = "StringValue"
                };
            var stringTemplate = new[]
                {
                    new[] {"", "Template:RootTemplate:A2:D4", "", ""},
                    new[] {"", "Адреса:", "Имена:", ""},
                    new[] {"", "Value::Array[].Address", "Value::Array[].Name", ""},
                    new[] {"Value::NonArray", "", "", ""}
                };

            var template = FakeTable.GenerateFromStringArray(stringTemplate);

            var target = new FakeTable(100, 100);
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger));
            var templateEngine = new TemplateEngine(template, logger);
            templateEngine.Render(tableBuilder, model);

            target.GetCell(new CellPosition("C2")).StringValue.Should().Be("Адреса:");
            target.GetCell(new CellPosition("D2")).StringValue.Should().Be("Имена:");
            target.GetCell(new CellPosition("C3")).StringValue.Should().Be("Address1");
            target.GetCell(new CellPosition("C4")).StringValue.Should().Be("Address2");
            target.GetCell(new CellPosition("C5")).StringValue.Should().BeEmpty();
            target.GetCell(new CellPosition("C6")).StringValue.Should().BeEmpty();
            target.GetCell(new CellPosition("C7")).StringValue.Should().Be("Address4");
            target.GetCell(new CellPosition("D3")).StringValue.Should().Be("Name1");
            target.GetCell(new CellPosition("D4")).StringValue.Should().BeEmpty();
            target.GetCell(new CellPosition("D5")).StringValue.Should().Be("Name3");
            target.GetCell(new CellPosition("D6")).StringValue.Should().BeEmpty();
            target.GetCell(new CellPosition("D7")).StringValue.Should().Be("Name4");
            target.GetCell(new CellPosition("B8")).StringValue.Should().Be("StringValue");

            DebugPrinting(target, new CellPosition(1, 1), new CellPosition(20, 20));
        }

        private static void DebugPrinting(ITable table, ICellPosition upperLeft, ICellPosition lowerRight)
        {
            var maxWidth = table.GetTablePart(new Rectangle(upperLeft, lowerRight))
                                .Cells
                                .SelectMany(row => row)
                                .Select(cell => cell.StringValue)
                                .Max(value => string.IsNullOrEmpty(value) ? 0 : value.Length);
            foreach (var row in table.GetTablePart(new Rectangle(upperLeft, lowerRight)).Cells)
            {
                foreach (var cell in row)
                {
                    var val = cell.StringValue + "";
                    while (val.Length < maxWidth)
                        val += " ";
                    Console.Write(val);
                }
                Console.WriteLine();
            }
        }

        private readonly ConsoleLog logger = new ConsoleLog();

        public class DocumentWithArray
        {
            public Organization[] Array { get; set; }
            public string NonArray { get; set; }
        }

        public class Document
        {
            public Organization DeliveryParty { get; set; }
            public Organization Shipper { get; set; }
            public Organization Invoicee { get; set; }
            public Organization Supplier { get; set; }
            public Organization Buyer { get; set; }
            public Organization Payer { get; set; }
            public string TypeName { get; set; }
            public VehicleInfo Vehicle { get; set; }
        }

        public class Organization
        {
            public string Name { get; set; }
            public string Inn { get; set; }
            public string Kpp { get; set; }
            public string Gln { get; set; }
            public string Address { get; set; }
            public string SupplierCodeInBuyerSystem { get; set; }
        }

        public class VehicleInfo
        {
            public string TransportMode { get; set; }
            public string VehicleNumber { get; set; }
            public string VehicleBrand { get; set; }
            public string NameOfCarrier { get; set; }
        }
    }
}