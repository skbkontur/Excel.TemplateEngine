using System;
using System.Collections;
using System.Collections.Generic;
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

        [TestCase("Value::StrProp", typeof(string), TestName = nameof(TypeExtractionTest) + " - string")]
        [TestCase("Value::IntProp", typeof(int), TestName = nameof(TypeExtractionTest) + " - int")]
        [TestCase("Value::DoubleProp", typeof(double), TestName = nameof(TypeExtractionTest) + " - double")]
        [TestCase("Value::ObjectProp", typeof(object), TestName = nameof(TypeExtractionTest) + " - null object")]
        [TestCase("Value::ArrayIntProp", typeof(int[]), TestName = nameof(TypeExtractionTest) + " - int array")]
        [TestCase("Value::NullableInt", typeof(int?), TestName = nameof(TypeExtractionTest) + " - nullable int")]
        [TestCase("Value::InnerObject.A.Value", typeof(MarkerA), TestName = nameof(TypeExtractionTest) + " - deep property")]
        [TestCase("Value::InnerObject.InnerArray[].ElementProp.A", typeof(MarkerB), TestName = nameof(TypeExtractionTest) + " - array in expression")]
        [TestCase("Value::InnerObject.InnerArray[1].ElementProp.A", typeof(MarkerB), TestName = nameof(TypeExtractionTest) + " - array with ind in expression")]
        [TestCase("Value::DictProperty", typeof(Dictionary<string, MarkerC>), TestName = nameof(TypeExtractionTest) + " - dict")]
        [TestCase("Value::DictProperty[\"Test\"]", typeof(MarkerC), TestName = nameof(TypeExtractionTest) + " - dict element")]
        [TestCase("Value::DictPropertyIntKeys[123]", typeof(MarkerD), TestName = nameof(TypeExtractionTest) + " - dict element int key")]
        public void TypeExtractionTest(string expression, Type expectedType)
        {
            var type = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelForTypeExtractionTest, ExpressionPath.FromRawExpression(expression));
            Assert.AreEqual(expectedType, type);
        }
        
        [TestCase("Value::lalala", TestName = "Non-existent property")]
        [TestCase("Value::InnerObject.A[].Value", TestName = "Array-access to non-array property")]
        [TestCase("Value::InnerObject[\"Test\"]", TestName = "Dict-access to non-dict property")]
        [TestCase("Value::DictProperty[]", TestName = "Access to dict with no index")]
        [TestCase("Value::DictProperty[123123]", TestName = "Access to dict with index of wrong type")]
        public void TypeExtractionTestOnInvalidExpressions(string expression)
        {
            Assert.Throws<ObjectPropertyExtractionException>(() => ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelForTypeExtractionTest, ExpressionPath.FromRawExpression(expression)));
        }

        [TestCase("Value::StrProp", typeof(string), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - string")]
        [TestCase("Value::IntProp", typeof(int), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - int")]
        [TestCase("Value::DoubleProp", typeof(double), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - double")]
        [TestCase("Value::ObjectProp", typeof(object), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - null object")]
        [TestCase("Value::ArrayIntProp", typeof(int[]), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - int array")]
        [TestCase("Value::NullableInt", typeof(int?), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - nullable int")]
        [TestCase("Value::InnerObject.A.Value", typeof(MarkerA), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - deep property")]
        [TestCase("Value::InnerObject.InnerArray[].ElementProp.A", typeof(MarkerB), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - array in expression")]
        [TestCase("Value::InnerObject.InnerArray[1].ElementProp.A", typeof(MarkerB), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - array with ind in expression")]
        [TestCase("Value::DictProperty", typeof(Dictionary<string, MarkerC>), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - dict")]
        [TestCase("Value::DictProperty[\"Test\"]", typeof(MarkerC), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - dict element")]
        [TestCase("Value::DictPropertyIntKeys[123]", typeof(MarkerD), TestName = nameof(TypeExtractionTestWithNotFullyInitializedModel) + " - dict element int key")]
        public void TypeExtractionTestWithNotFullyInitializedModel(string expression, Type expectedType) // todo (mpivko, 19.01.2018): merge with TypeExtractionTest
        {
            var modelToGetType = new ComplexModel();
            
            var type = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelToGetType, ExpressionPath.FromRawExpression(expression));
            Assert.AreEqual(expectedType, type);
        }

        private static object[][] ExtractChildObjectSetterTestSource { [UsedImplicitly] get; } = new (string, Func<ComplexModel, object>, object)[]
            {
                ("Value::StrProp", x => x.StrProp, "123"),
                ("Value::IntProp", x => x.IntProp, 42),
                ("Value::DoubleProp", x => x.DoubleProp, 1.5),
                ("Value::ObjectProp", x => x.ObjectProp, new MarkerA()),
                ("Value::ArrayIntProp", x => x.ArrayIntProp, new[] {1, 2, 10}),
                ("Value::NullableInt", x => x.NullableInt, (int?)10),
                ("Value::InnerObject.A.Value", x => x.InnerObject.A.Value, new MarkerA()),
                ("Value::InnerObject.InnerArray[1].ElementProp.A", x => x.InnerObject.InnerArray[1].ElementProp.A, new MarkerB()),// todo mpivko
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
            var s = ObjectPropertiesExtractor.ExtractChildObjectSetter(modelToSet, expression);
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

        public class MarkerA { }
        public class MarkerB { }
        public class MarkerC { }
        public class MarkerD { }


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
        }

        public class C
        {
            public string S { get; set; }
        }

        #endregion
    }
}