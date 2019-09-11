using System;
using System.Linq;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.FakeDocumentPrimitivesImplementation;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using Excel.TemplateEngine.ObjectPrinting.TableNavigator;

using NUnit.Framework;

using Vostok.Logging.Console;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
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

            Assert.AreEqual("TestStringValue", target.GetCell(new CellPosition("B2")).StringValue);

            Assert.AreEqual("(1,1)", ((FakeCell)target.GetCell(new CellPosition("B2"))).StyleId);
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

            Assert.AreEqual("a", target.GetCell(new CellPosition("B2")).StringValue);
            Assert.AreEqual("b", target.GetCell(new CellPosition("B3")).StringValue);
            Assert.AreEqual("c", target.GetCell(new CellPosition("B4")).StringValue);
            Assert.AreEqual("d", target.GetCell(new CellPosition("B5")).StringValue);
            Assert.AreEqual("e", target.GetCell(new CellPosition("B6")).StringValue);

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

            Assert.AreEqual("Покупатель:", target.GetCell(new CellPosition("A1")).StringValue);
            Assert.AreEqual("Поставщик:", target.GetCell(new CellPosition("B1")).StringValue);
            Assert.AreEqual("BuyerName", target.GetCell(new CellPosition("A2")).StringValue);
            Assert.AreEqual("SupplierName", target.GetCell(new CellPosition("B2")).StringValue);
            Assert.AreEqual("ORDERS", target.GetCell(new CellPosition("C2")).StringValue);
            Assert.AreEqual("BuyerAddress", target.GetCell(new CellPosition("A3")).StringValue);
            Assert.AreEqual("SupplierAddress", target.GetCell(new CellPosition("B3")).StringValue);

            Assert.AreEqual("B4", ((FakeCell)target.GetCell(new CellPosition("A1"))).StyleId);
            Assert.AreEqual("C4", ((FakeCell)target.GetCell(new CellPosition("B1"))).StyleId);
            Assert.AreEqual("A8", ((FakeCell)target.GetCell(new CellPosition("A2"))).StyleId);
            Assert.AreEqual("A8", ((FakeCell)target.GetCell(new CellPosition("B2"))).StyleId);
            Assert.AreEqual("D5", ((FakeCell)target.GetCell(new CellPosition("C2"))).StyleId);
            Assert.AreEqual("A9", ((FakeCell)target.GetCell(new CellPosition("A3"))).StyleId);
            Assert.AreEqual("A9", ((FakeCell)target.GetCell(new CellPosition("B3"))).StyleId);

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

            Assert.AreEqual("Покупатель:", target.GetCell(new CellPosition("A1")).StringValue);
            Assert.AreEqual("Поставщик:", target.GetCell(new CellPosition("C1")).StringValue);
            Assert.AreEqual("BuyerName", target.GetCell(new CellPosition("A2")).StringValue);
            Assert.AreEqual("BuyerAddress", target.GetCell(new CellPosition("A3")).StringValue);
            Assert.AreEqual("ORDERS", target.GetCell(new CellPosition("B2")).StringValue);
            Assert.AreEqual("SupplierName", target.GetCell(new CellPosition("C2")).StringValue);
            Assert.AreEqual("90238192038", target.GetCell(new CellPosition("D2")).StringValue);
            Assert.AreEqual("SupplierAddress", target.GetCell(new CellPosition("C3")).StringValue);
            Assert.AreEqual("0832309812", target.GetCell(new CellPosition("D3")).StringValue);
            Assert.AreEqual("Плательщик:", target.GetCell(new CellPosition("A4")).StringValue);
            Assert.AreEqual("Средство доставки:", target.GetCell(new CellPosition("B4")).StringValue);
            Assert.AreEqual("Доставка:", target.GetCell(new CellPosition("D4")).StringValue);
            Assert.AreEqual("PayerName", target.GetCell(new CellPosition("A5")).StringValue);
            Assert.AreEqual("PayerAddress", target.GetCell(new CellPosition("A6")).StringValue);
            Assert.AreEqual("Евкакий", target.GetCell(new CellPosition("B5")).StringValue);
            Assert.AreEqual("Tauria", target.GetCell(new CellPosition("C5")).StringValue);
            Assert.AreEqual("Agressive", target.GetCell(new CellPosition("B6")).StringValue);
            Assert.AreEqual("A777AB", target.GetCell(new CellPosition("C6")).StringValue);
            Assert.AreEqual("DeliveryPartyName", target.GetCell(new CellPosition("D5")).StringValue);
            Assert.AreEqual("DeliveryPartyAddress", target.GetCell(new CellPosition("D6")).StringValue);

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

            Assert.AreEqual("Адреса:", target.GetCell(new CellPosition("C2")).StringValue);
            Assert.AreEqual("Имена:", target.GetCell(new CellPosition("D2")).StringValue);
            Assert.AreEqual("Address1", target.GetCell(new CellPosition("C3")).StringValue);
            Assert.AreEqual("Address2", target.GetCell(new CellPosition("C4")).StringValue);
            Assert.AreEqual("Address3", target.GetCell(new CellPosition("C5")).StringValue);
            Assert.AreEqual("Address4", target.GetCell(new CellPosition("C6")).StringValue);
            Assert.AreEqual("Name1", target.GetCell(new CellPosition("D3")).StringValue);
            Assert.AreEqual("Name2", target.GetCell(new CellPosition("D4")).StringValue);
            Assert.AreEqual("Name3", target.GetCell(new CellPosition("D5")).StringValue);
            Assert.AreEqual("Name4", target.GetCell(new CellPosition("D6")).StringValue);
            Assert.AreEqual("StringValue", target.GetCell(new CellPosition("B7")).StringValue);

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

            Assert.AreEqual("Адреса:", target.GetCell(new CellPosition("C2")).StringValue);
            Assert.AreEqual("Имена:", target.GetCell(new CellPosition("D2")).StringValue);
            Assert.AreEqual("", target.GetCell(new CellPosition("C3")).StringValue);
            Assert.AreEqual("", target.GetCell(new CellPosition("D3")).StringValue);
            Assert.AreEqual("StringValue", target.GetCell(new CellPosition("B4")).StringValue);

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

            Assert.AreEqual("Адреса:", target.GetCell(new CellPosition("C2")).StringValue);
            Assert.AreEqual("Имена:", target.GetCell(new CellPosition("D2")).StringValue);
            Assert.AreEqual("Address1", target.GetCell(new CellPosition("C3")).StringValue);
            Assert.AreEqual("Address2", target.GetCell(new CellPosition("C4")).StringValue);
            Assert.AreEqual("", target.GetCell(new CellPosition("C5")).StringValue);
            Assert.AreEqual("", target.GetCell(new CellPosition("C6")).StringValue);
            Assert.AreEqual("Address4", target.GetCell(new CellPosition("C7")).StringValue);
            Assert.AreEqual("Name1", target.GetCell(new CellPosition("D3")).StringValue);
            Assert.AreEqual("", target.GetCell(new CellPosition("D4")).StringValue);
            Assert.AreEqual("Name3", target.GetCell(new CellPosition("D5")).StringValue);
            Assert.AreEqual("", target.GetCell(new CellPosition("D6")).StringValue);
            Assert.AreEqual("Name4", target.GetCell(new CellPosition("D7")).StringValue);
            Assert.AreEqual("StringValue", target.GetCell(new CellPosition("B8")).StringValue);

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
                        val = val + " ";
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