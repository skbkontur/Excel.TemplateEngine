using System;
using System.Collections;
using System.Collections.Generic;
using System.Diagnostics.CodeAnalysis;
using System.Linq;

using JetBrains.Annotations;

using NUnit.Framework;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;

namespace SKBKontur.Catalogue.Core.Tests.ExcelObjectPrinterTests
{
    [TestFixture]
    public class ObjectPropertiesExtractorTests
    {
        [Test]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void AtomicObjectsArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[].Cs[].S";

            var child = ObjectPropertiesExtractor.ExtractChildObject(model, ExcelTemplatePath.FromRawExpression(valueDesription));
            var childArray = child as object[];
            Assert.NotNull(childArray);

            Assert.AreEqual(2, childArray.Length);
            CollectionAssert.AreEqual(model.Bs[0].Cs.Select(x => x.S), childArray[0] as object[]);
            CollectionAssert.AreEqual(model.Bs[1].Cs.Select(x => x.S), childArray[1] as object[]);
        }

        [Test]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void ComplexObjectsArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[].Cs[]";
            var child = ObjectPropertiesExtractor.ExtractChildObject(model, ExcelTemplatePath.FromRawExpression(valueDesription));
            var childArray = child as object[];
            Assert.NotNull(childArray);

            Assert.AreEqual(2, childArray.Length);
            CollectionAssert.AreEqual(model.Bs[0].Cs, childArray[0] as object[]);
            CollectionAssert.AreEqual(model.Bs[1].Cs, childArray[1] as object[]);
        }

        [Test]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void NullArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs";
            var child = ObjectPropertiesExtractor.ExtractChildObject(new A(), ExcelTemplatePath.FromRawExpression(valueDesription));
            Assert.Null(child);
        }

        [Test]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void NullArrayExtractionTestWithBraces()
        {
            const string valueDesription = "Value::Bs[]";
            var child = ObjectPropertiesExtractor.ExtractChildObject(new A(), ExcelTemplatePath.FromRawExpression(valueDesription));
            Assert.Null(child);
        }

        [Test]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void NullValuesArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[]";
            var localModel = new A
                {
                    Bs = new[] {null, new B(), null}
                };
            var child = ObjectPropertiesExtractor.ExtractChildObject(localModel, ExcelTemplatePath.FromRawExpression(valueDesription));
            var childArray = child as object[];
            CollectionAssert.AreEqual(localModel.Bs, childArray);
        }

        [Test]
        [SuppressMessage("ReSharper", "PossibleNullReferenceException")]
        public void AllNullValuesArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[]";
            var localModel = new A
                {
                    Bs = new B[] {null, null, null}
                };
            var child = ObjectPropertiesExtractor.ExtractChildObject(localModel, ExcelTemplatePath.FromRawExpression(valueDesription));
            var childArray = child as object[];
            CollectionAssert.AreEqual(localModel.Bs, childArray);
        }

        public static IEnumerable CollectionElementsExtractionTestSource
        {
            get
            {
                return new (string, Func<ModelWithCollections, object>)[]
                    {
                        ("IntToStringDict[10]", x => x.IntToStringDict[10]),
                        ("IntToStringDict[25]", x => x.IntToStringDict[25]),
                        ("StringToIntDict[\"lalala\"]", x => x.StringToIntDict["lalala"]),
                        ("StringToIntDict[\"abracadabra\"]", x => x.StringToIntDict["abracadabra"]),
                        ("Array[0]", x => x.Array[0]),
                        ("Array[1]", x => x.Array[1]),
                        ("Array[2]", x => x.Array[2]),
                        ("List[0]", x => x.List[0]),
                        ("List[1]", x => x.List[1]),
                        ("List[2]", x => x.List[2]),
                        ("InnerModel.IntToStringDict[10010]", x => x.InnerModel.IntToStringDict[10010]),
                        ("InnerModel.IntToStringDict[10025]", x => x.InnerModel.IntToStringDict[10025]),
                        ("InnerModel.StringToIntDict[\"inner_lalala\"]", x => x.InnerModel.StringToIntDict["inner_lalala"]),
                        ("InnerModel.StringToIntDict[\"inner_abracadabra\"]", x => x.InnerModel.StringToIntDict["inner_abracadabra"]),
                        ("InnerModel.Array[0]", x => x.InnerModel.Array[0]),
                        ("InnerModel.Array[1]", x => x.InnerModel.Array[1]),
                        ("InnerModel.Array[2]", x => x.InnerModel.Array[2]),
                        ("InnerModel.List[0]", x => x.InnerModel.List[0]),
                        ("InnerModel.List[1]", x => x.InnerModel.List[1]),
                        ("InnerModel.List[2]", x => x.InnerModel.List[2]),
                    }
                    .Select(x => new object[] {$"Value::{x.Item1}", x.Item2});
            }
        }

        [TestCaseSource(nameof(CollectionElementsExtractionTestSource))]
        public void CollectionElementsExtractionTest(string valueDesription, Func<ModelWithCollections, object> expectedElementProvider)
        {
            var localModel = new ModelWithCollections
                {
                    IntToStringDict = new Dictionary<int, string> {{10, "abc"}, {25, "def"}},
                    StringToIntDict = new Dictionary<string, int> {{"lalala", 123}, {"abracadabra", 321}},
                    Array = new[] {"a", "lalaa", "test"},
                    List = new List<string> {"first", "second", "third"},
                    InnerModel = new InnerModelWithCollections
                        {
                            IntToStringDict = new Dictionary<int, string> {{10010, "inner_abc"}, {10025, "inner_def"}},
                            StringToIntDict = new Dictionary<string, int> {{"inner_lalala", 100123}, {"inner_abracadabra", 100321}},
                            Array = new[] {"inner_a", "inner_inner_lalaa", "inner_test"},
                            List = new List<string> {"first", "inner_second", "inner_third"},
                        }
                };
            var child = ObjectPropertiesExtractor.ExtractChildObject(localModel, ExcelTemplatePath.FromRawExpression(valueDesription));
            Assert.AreEqual(expectedElementProvider(localModel), child);
        }

        [Test]
        public void NonexistentObjectsArrayExtractionTest()
        {
            const string valueDesription = "Value::Bs[].Cs[].NULL";
            Assert.Throws<ObjectPropertyExtractionException>(() => ObjectPropertiesExtractor.ExtractChildObject(model, ExcelTemplatePath.FromRawExpression(valueDesription)));
        }

        [Test]
        public void NonexistentObjectExtractionTest()
        {
            const string valueDesription = "Value::NULL";
            Assert.Throws<ObjectPropertyExtractionException>(() => ObjectPropertiesExtractor.ExtractChildObject(model, ExcelTemplatePath.FromRawExpression(valueDesription)));
        }

        [Test]
        public void AtomicValueExtractionTest()
        {
            var simpleModel = new C
                {
                    S = "Test"
                };
            const string valueDescription = "Value::S";
            var child = ObjectPropertiesExtractor.ExtractChildObject(simpleModel, ExcelTemplatePath.FromRawExpression(valueDescription));
            Assert.AreNotEqual(null, child);
            Assert.AreEqual("Test", child);
        }

        [TestCase("Value::lalala", TestName = "Non-existent property")]
        [TestCase("Value::InnerObject.A[].Value", TestName = "Array-access to non-array property")]
        [TestCase("Value::InnerObject[\"Test\"]", TestName = "Dict-access to non-dict property")]
        [TestCase("Value::DictProperty[]", TestName = "Access to dict with no index")]
        [TestCase("Value::DictProperty[123123]", TestName = "Access to dict with index of wrong type")]
        public void TypeExtractionTestOnInvalidExpressions(string expression)
        {
            Assert.Throws<ObjectPropertyExtractionException>(() => ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelForTypeExtractionTest.GetType(), ExcelTemplatePath.FromRawExpression(expression)));
        }

        public static IEnumerable TypeExtractionTestTestCases
        {
            get
            {
                yield return new TestCaseData("Value::IntProp", typeof(int)).SetName("int");
                yield return new TestCaseData("Value::StrProp", typeof(string)).SetName("string");
                yield return new TestCaseData("Value::DoubleProp", typeof(double)).SetName("double");
                yield return new TestCaseData("Value::ObjectProp", typeof(object)).SetName("null object");
                yield return new TestCaseData("Value::ArrayIntProp", typeof(int[])).SetName("int array");
                yield return new TestCaseData("Value::NullableInt", typeof(int?)).SetName("nullable int");
                yield return new TestCaseData("Value::InnerObject.A.Value", typeof(MarkerA)).SetName("deep property");
                yield return new TestCaseData("Value::InnerObject.InnerArray[].ElementProp.A", typeof(MarkerB)).SetName("array in expression");
                yield return new TestCaseData("Value::InnerObject.InnerArray[1].ElementProp.A", typeof(MarkerB)).SetName("array with ind in expression");
                yield return new TestCaseData("Value::DictProperty", typeof(Dictionary<string, MarkerC>)).SetName("dict");
                yield return new TestCaseData("Value::DictProperty[\"Test\"]", typeof(MarkerC)).SetName("dict element");
                yield return new TestCaseData("Value::DictPropertyIntKeys[123]", typeof(MarkerD)).SetName("dict element int key");
            }
        }

        [TestCaseSource(nameof(TypeExtractionTestTestCases))]
        public void TypeExtractionTest(string expression, Type expectedType)
        {
            var type = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelForTypeExtractionTest.GetType(), ExcelTemplatePath.FromRawExpression(expression));
            Assert.AreEqual(expectedType, type);
        }

        [TestCaseSource(nameof(TypeExtractionTestTestCases))]
        public void TypeExtractionTestWithNotFullyInitializedModel(string expression, Type expectedType)
        {
            var modelToGetType = new ComplexModel();

            var type = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelToGetType.GetType(), ExcelTemplatePath.FromRawExpression(expression));
            Assert.AreEqual(expectedType, type);
        }

        private static object[][] ExtractChildObjectSetterTestSource { [UsedImplicitly] get; } = new (string, Func<ComplexModel, object>, object)[]
            {
                ("Value::StrProp", x => x.StrProp, "123"),
                ("Value::StrProp", x => x.StrProp, null),
                ("Value::IntProp", x => x.IntProp, 42),
                ("Value::DoubleProp", x => x.DoubleProp, 1.5),
                ("Value::ObjectProp", x => x.ObjectProp, new MarkerA()),
                ("Value::ArrayIntProp", x => x.ArrayIntProp, new[] {1, 2, 10}),
                ("Value::NullableInt", x => x.NullableInt, (int?)10),
                ("Value::InnerObject.A.Value", x => x.InnerObject.A.Value, new MarkerA()),
                ("Value::InnerObject.InnerArray[1].ElementProp.A", x => x.InnerObject.InnerArray[1].ElementProp.A, new MarkerB()),
                ("Value::DictProperty", x => x.DictProperty, new Dictionary<string, MarkerC>()),
                ("Value::DictProperty[\"Test\"]", x => x.DictProperty["Test"], new MarkerC()),
                ("Value::DictPropertyIntKeys[123]", x => x.DictPropertyIntKeys[123], new MarkerD()),
                ("Value::DictPropertyClasses[\"Lalala\"].A.Value", x => x.DictPropertyClasses["Lalala"].A.Value, new MarkerA()),
                ("Value::ArrayOfObjects[].A.SimpleValue", x => x.ArrayOfObjects.Select(y => y.A.SimpleValue).ToArray(), new[] {1, 2, 10}),
                ("Value::ArrayOfObjects[].A.SimpleValue", x => x.ArrayOfObjects.Select(y => y.A.SimpleValue).ToArray(), new List<int> {1, 2, 10}.Cast<object>().ToList()),
            }
            .Select(x => new[] {x.Item1, x.Item2, x.Item3}).ToArray();

        [TestCaseSource(nameof(ExtractChildObjectSetterTestSource))]
        public void ExtractChildObjectSetterTest(string expression, Func<ComplexModel, object> getter, object valueToSet)
        {
            var modelToSet = new ComplexModel();
            var s = ObjectPropertySettersExtractor.ExtractChildObjectSetter(modelToSet, ExcelTemplatePath.FromRawExpression(expression));
            s(valueToSet);
            Assert.AreEqual(valueToSet, getter(modelToSet));
        }

        #region auxiliary models

        public class ComplexModel
        {
            public string StrProp { get; set; }
            public int IntProp { get; set; }
            public double DoubleProp { get; set; }
            public object ObjectProp { get; set; }
            public int[] ArrayIntProp { get; set; }
            public int? NullableInt { get; set; }
            public InnerClass InnerObject { get; set; }
            public Dictionary<string, MarkerC> DictProperty { get; set; }
            public Dictionary<int, MarkerD> DictPropertyIntKeys { get; set; }
            public Dictionary<string, InnerClass> DictPropertyClasses { get; set; }
            public InnerClass[] ArrayOfObjects { get; set; }

            public class InnerClass
            {
                public SubInnerClass A { get; set; }
                public SubInnerArrayElement[] InnerArray { get; set; }

                public class SubInnerClass
                {
                    public MarkerA Value { get; set; }
                    public int SimpleValue { get; set; }
                }

                public class SubInnerArrayElement
                {
                    public ElementInnerClass ElementProp { get; set; }

                    public class ElementInnerClass
                    {
                        public MarkerB A { get; set; }
                    }
                }
            }
        }

        private readonly object modelForTypeExtractionTest = new
            {
                StrProp = "some str",
                IntProp = 123,
                DoubleProp = 1.5,
                ObjectProp = (object)null,
                ArrayIntProp = new int[10],
                NullableInt = (int?)123,
                InnerObject = new
                    {
                        A = new
                            {
                                Value = new MarkerA()
                            },
                        InnerArray = new[]
                            {
                                new
                                    {
                                        ElementProp = new
                                            {
                                                A = new MarkerB()
                                            }
                                    },
                                new
                                    {
                                        ElementProp = new
                                            {
                                                A = new MarkerB()
                                            }
                                    },
                            },
                    },
                DictProperty = new Dictionary<string, MarkerC>
                    {
                        {"Test", new MarkerC()},
                    },
                DictPropertyIntKeys = new Dictionary<int, MarkerD>
                    {
                        {123, new MarkerD()},
                    },
            };

        public class MarkerA
        {
        }

        public class MarkerB
        {
        }

        public class MarkerC
        {
        }

        public class MarkerD
        {
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

        public class A
        {
            public B[] Bs { get; set; }
        }

        public class B
        {
            public C[] Cs { get; set; }

            protected bool Equals(B other)
            {
                return Equals(Cs, other.Cs);
            }

            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj)) return false;
                if (ReferenceEquals(this, obj)) return true;
                if (obj.GetType() != this.GetType()) return false;
                return Equals((B)obj);
            }

            public override int GetHashCode()
            {
                return (Cs != null ? Cs.GetHashCode() : 0);
            }
        }

        public class C
        {
            public string S { get; set; }

            protected bool Equals(C other)
            {
                return string.Equals(S, other.S);
            }

            public override bool Equals(object obj)
            {
                if (ReferenceEquals(null, obj)) return false;
                if (ReferenceEquals(this, obj)) return true;
                if (obj.GetType() != this.GetType()) return false;
                return Equals((C)obj);
            }

            public override int GetHashCode()
            {
                return (S != null ? S.GetHashCode() : 0);
            }
        }

        public class ModelWithCollections
        {
            public Dictionary<string, int> StringToIntDict { get; set; }
            public Dictionary<int, string> IntToStringDict { get; set; }
            public string[] Array { get; set; }
            public List<string> List { get; set; }
            public InnerModelWithCollections InnerModel { get; set; }
        }

        public class InnerModelWithCollections
        {
            public Dictionary<string, int> StringToIntDict { get; set; }
            public Dictionary<int, string> IntToStringDict { get; set; }
            public string[] Array { get; set; }
            public List<string> List { get; set; }
        }

        #endregion
    }
}