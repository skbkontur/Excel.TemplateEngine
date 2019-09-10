using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

using Excel.TemplateEngine.FileGenerating;
using Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitivesImplementation;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using Excel.TemplateEngine.ObjectPrinting.TableNavigator;

using NUnit.Framework;

using Vostok.Logging.Console;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests
{
    [TestFixture]
    public class ObjectExcelPrintingTests : FileBasedTestBase
    {
        [Test]
        public void PrintStringTest()
        {
            MakeTest("Test text", "template.xlsx");
        }

        [Test]
        public void PrintSimpleObjectTest()
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

            MakeTest(model, "template.xlsx");
        }

        [Test]
        public void PrintWithCellsMergingTest()
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

            MakeTest(model, "withCellsMergingTemplate.xlsx", doc =>
                {
                    var mergedCells = doc.MergedCells.FirstOrDefault();
                    Assert.AreNotEqual(null, mergedCells);
                    Assert.AreEqual("C2", mergedCells.UpperLeft.CellReference);
                    Assert.AreEqual("D2", mergedCells.LowerRight.CellReference);
                });
        }

        [Test]
        public void PrintObjectsByFullPathTest()
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

            MakeTest(model, "withFullPathsTemplate.xlsx");
        }

        [Test]
        public void PrintComplexObjectTest()
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
                            Inn = 90238192,
                            Kpp = "0832309812"
                        },
                    Payer = new Organization
                        {
                            Address = "PayerAddress",
                            Name = "PayerAddress"
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

            MakeTest(model, "complexTemplate.xlsx");
        }

        [Test]
        public void MultipleObjectsPrintingTest()
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
                            Inn = 90238192,
                            Kpp = "0832309812"
                        },
                    Payer = new Organization
                        {
                            Address = "PayerAddress",
                            Name = "PayerAddress"
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

            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("complexTemplate.xlsx")), logger);
            var template = new ExcelTable(templateDocument.GetWorksheet(0));
            var templateEngine = new TemplateEngine(template, logger);

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), logger);
            targetDocument.AddWorksheet("Лист2");

            var target = new ExcelTable(targetDocument.GetWorksheet(0));
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger), new Style(template.GetCell(new CellPosition("A1"))));
            templateEngine.Render(tableBuilder, model);

            target = new ExcelTable(targetDocument.GetWorksheet(1));
            tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("A1"), logger), new Style(template.GetCell(new CellPosition("A1"))));
            templateEngine.Render(tableBuilder, model);

            var result = targetDocument.CloseAndGetDocumentBytes();
            File.WriteAllBytes("output.xlsx", result);

            templateDocument.Dispose();
            targetDocument.Dispose();
        }

        [Test]
        public void FormControlsPrintingTest()
        {
            var model = new
                {
                    BoolValue1 = false,
                    BoolValue2 = true,
                    String1 = "Value1",
                    Dict = new Dictionary<string, bool> {{"TestKey", false}},
                };

            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("formControlsTemplate.xlsx")), logger);
            var template = new ExcelTable(templateDocument.GetWorksheet(0));
            var templateEngine = new TemplateEngine(template, logger);

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), logger);

            var target = new ExcelTable(targetDocument.GetWorksheet(0));
            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("A1"), logger), new Style(template.GetCell(new CellPosition("A1"))));
            templateEngine.Render(tableBuilder, model);

            var result = targetDocument.CloseAndGetDocumentBytes();
            File.WriteAllBytes("output.xlsx", result);

            templateDocument.Dispose();
            targetDocument.Dispose();
        }

        [Test]
        public void NonExistentFieldPrintingTest()
        {
            var model = new
                {
                    A = false,
                    B = true,
                };

            using (var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("template.xlsx")), logger))
            {
                var template = new ExcelTable(templateDocument.GetWorksheet(0));
                var templateEngine = new TemplateEngine(template, logger);

                using (var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), logger))
                {
                    var target = new ExcelTable(targetDocument.GetWorksheet(0));
                    var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("A1"), logger), new Style(template.GetCell(new CellPosition("A1"))));

                    Assert.Throws<ExcelEngineException>(() => templateEngine.Render(tableBuilder, model));
                }
            }
        }

        private void MakeTest(object model, string templateFileName, Action<ExcelTable> resultValidationFunc = null)
        {
            var templateDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath(templateFileName)), logger);
            var template = new ExcelTable(templateDocument.GetWorksheet(0));

            var targetDocument = ExcelDocumentFactory.CreateFromTemplate(File.ReadAllBytes(GetFilePath("empty.xlsx")), logger);
            var target = new ExcelTable(targetDocument.GetWorksheet(0));

            var tableBuilder = new TableBuilder(target, new TableNavigator(new CellPosition("B2"), logger), new Style(template.GetCell(new CellPosition("A1"))));
            var templateEngine = new TemplateEngine(template, logger);
            templateEngine.Render(tableBuilder, model);

            var result = targetDocument.CloseAndGetDocumentBytes();
            File.WriteAllBytes("output.xlsx", result);

            resultValidationFunc?.Invoke(target);

            templateDocument.Dispose();
            targetDocument.Dispose();
        }

        private readonly ConsoleLog logger = new ConsoleLog();

        public class DocumentWithArray
        {
            public Organization[][] Array { get; set; }
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
            public int Inn { get; set; }
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