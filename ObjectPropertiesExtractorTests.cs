using System.Collections;
using System.Linq;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class ObjectPropertiesExtractorTests
    {
        [Test]
        public void AtomicObjectsArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[].Cs[].S";

            var child = ObjectPropertiesExtractor.Instance.ExtractChildObject(model, valueDesription);
            Assert.AreNotEqual(null, child);

            var childArray = ((IEnumerable)child).Cast<string>().ToArray();
            Assert.AreNotEqual(null, childArray);
// ReSharper disable AssignNullToNotNullAttribute
            Assert.AreEqual(6, childArray.Count());
// ReSharper restore AssignNullToNotNullAttribute
// ReSharper disable PossibleNullReferenceException
            Assert.AreEqual("Test21", childArray[3]);
// ReSharper restore PossibleNullReferenceException
        }

        [Test]
        public void ComplexObjectsArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[].Cs[]";
            var child = ObjectPropertiesExtractor.Instance.ExtractChildObject(model, valueDesription);
            Assert.AreNotEqual(null, child);
            var objectChildArray = ((IEnumerable)child).Cast<C>().ToArray();
            Assert.AreNotEqual(null, objectChildArray);
            Assert.AreEqual(6, objectChildArray.Count());
        }

        [Test]
        public void NonexistentObjectsArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[].Cs[].NULL";
            var child = ObjectPropertiesExtractor.Instance.ExtractChildObject(model, valueDesription);
            Assert.AreEqual(null, child);
        }

        [Test]
        public void NonexistentObjectExtractionTest()
        {
            const string valueDesription = "Value::NULL";
            var child = ObjectPropertiesExtractor.Instance.ExtractChildObject(model, valueDesription);
            Assert.AreEqual(null, child);
        }

        [Test]
        public void AtomicValueExtractionTest()
        {
            var simpleModel = new C
                {
                    S = "Test"
                };
            const string valueDescription = "Value::S";
            var child = ObjectPropertiesExtractor.Instance.ExtractChildObject(simpleModel, valueDescription);
            Assert.AreNotEqual(null, child);
            Assert.AreEqual("Test", child);
        }

        public class A
        {
            public B[] Bs { get; set; }
        }

        public class B
        {
            public C[] Cs { get; set; }
        }

        public class C
        {
            public string S { get; set; }
        }

        private readonly A model = new A
            {
                Bs = new[]
                    {
                        new B
                            {
                                Cs = new[]
                                    {
                                        new C
                                            {
                                                S = "Test11"
                                            },
                                        new C
                                            {
                                                S = "Test12"
                                            },
                                        new C
                                            {
                                                S = "Test13"
                                            }
                                    }
                            },
                        new B
                            {
                                Cs = new[]
                                    {
                                        new C
                                            {
                                                S = "Test21"
                                            },
                                        new C
                                            {
                                                S = "Test22"
                                            },
                                        new C
                                            {
                                                S = null
                                            }
                                    }
                            }
                    }
            };
    }
}