using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.Linq;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public sealed class ObjectPropertiesExtractor
    {
        private ObjectPropertiesExtractor()
        {
        }

        [CanBeNull]
        public object ExtractChildObject(object model, string expression)
        {
            if(!(TemplateDescriptionHelper.Instance.IsCorrectValueDescription(expression) || TemplateDescriptionHelper.Instance.IsCorrectFormValueDescription(expression)) ||
               model == null)
                return null;

            var descriptionParts = TemplateDescriptionHelper.Instance.GetDescriptionParts(expression);
            var pathParts = descriptionParts[2].Split('.');

            return ExtractChildObject(model, pathParts);
        }

        [CanBeNull]
        public object ExtractChildObjectViaPath(object model, string path)
        {
            if (!TemplateDescriptionHelper.Instance.IsCorrectModelPath(path))
                return null;
            var pathParts = path.Split('.');
            return ExtractChildObject(model, pathParts);
        }

        [NotNull]
        public static Type ExtractChildObjectTypeFromPath([NotNull] object model/*todo mpivko: we need only type, not model*/, [NotNull] ExpressionPath path) // todo (mpivko, 19.01.2018): make sure that path is not cleaned
        {
            var currType = model.GetType();

            foreach (var part in path.PartsWithIndexers)
            {
                var childPropertyType = ExtractPropertyInfo(currType, part).PropertyType;

                if (TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(part))
                {
                    if(TypeCheckingHelper.Instance.IsDictionary(childPropertyType))
                    {
                        ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), TypeCheckingHelper.Instance.GetDictionaryKeyType(childPropertyType));
                        currType = TypeCheckingHelper.Instance.GetDictionaryValueType(childPropertyType);
                    }
                    else if(TypeCheckingHelper.Instance.IsIList(childPropertyType))
                    {
                        ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), typeof(int));
                        currType = TypeCheckingHelper.Instance.GetEnumerableItemType(childPropertyType);
                    }
                    else
                        throw new ObjectPropertyExtractionException($"Not supported collection type {childPropertyType}");
                }
                else if(TemplateDescriptionHelper.Instance.IsArrayPathPart(part))
                {
                    if (TypeCheckingHelper.Instance.IsIList(childPropertyType))
                        currType = TypeCheckingHelper.Instance.GetIListItemType(childPropertyType);
                    else
                        throw new ObjectPropertyExtractionException($"Not supported collection type {childPropertyType}");
                }
                else
                {
                    currType = childPropertyType;
                }
            }
            return currType;
        }

        public static (string, string) SplitForEnumerableExpansion([NotNull] string expression)
        {
            var parts = ExtractChildObjectPath(expression).Split('.');
            if (!parts.Any(x => x.Contains("[]")))
                throw new ObjectPropertyExtractionException($"Expression needs enumerable expansion but has no part with '[]' (expression - '{expression}')");
            var firstPartLen = parts.TakeWhile(x => !x.EndsWith("[]")).Count() + 1;
            return (string.Join(".", parts.Take(firstPartLen)), string.Join(".", parts.Skip(firstPartLen)));
        }

        public static bool NeedEnumerableExpansion([NotNull] string expression)
        {
            return ExtractChildObjectPath(expression).Contains("[]");
        }

        public static string ExtractChildObjectPath([NotNull] string expression)
        {
            // todo (mpivko, 19.01.2018): check expression
            return expression.Split(':')[2];
        }

        public static string ExtractCleanChildObjectPath([NotNull] string expression)
        {
            return ExtractChildObjectPath(expression).Replace("[]", "");
        }

        [NotNull]
        public static Action<object> ExtractChildObjectSetter([NotNull] object model, [NotNull] string expression)
        {
            if (!TemplateDescriptionHelper.Instance.IsCorrectAbstractValueDescription(expression))
                throw new ObjectPropertyExtractionException($"Invalid description {expression}");

            var descriptionParts = TemplateDescriptionHelper.Instance.GetDescriptionParts(expression);
            var pathParts = descriptionParts[2].Split('.');

            return ExtractChildModelSetter(model, pathParts);
        }
        
        public static ObjectPropertiesExtractor Instance { get; } = new ObjectPropertiesExtractor();

        private static bool TryExtractCurrentChildPropertyInfo(object model, string pathPart, out PropertyInfo childPropertyInfo)
        {
            var propertyName = TemplateDescriptionHelper.Instance.GetPathPartName(pathPart);
            childPropertyInfo = model.GetType().GetProperty(propertyName);
            return childPropertyInfo != null;
        }

        [NotNull]
        private static PropertyInfo ExtractPropertyInfo(Type type, string pathPart)
        {
            var propertyName = TemplateDescriptionHelper.Instance.GetPathPartName(pathPart);
            var childPropertyInfo = type.GetProperty(propertyName);
            if (childPropertyInfo == null)
                throw new ObjectPropertyExtractionException($"Property with name '{propertyName}' not found in type '{type}'");
            return childPropertyInfo;
        }

        private static bool TryExtractCurrentChild(object model, string pathPart, out object child)
        {
            child = null;
            if (TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(pathPart))
            {
                var name = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartName(pathPart);
                var key = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(pathPart);
                if (!TryExtractCurrentChildPropertyInfo(model, name, out var dictPropertyInfo))
                    return false;
                if(!TypeCheckingHelper.Instance.IsDictionary(dictPropertyInfo.PropertyType))
                    throw new ObjectPropertyExtractionException($"Unexpected child type: expected dictionary (pathPath='{pathPart}'), but model is '{dictPropertyInfo.PropertyType}' in '{model.GetType()}'"); // todo (mpivko, 30.01.2018): what about arrays?
                var indexer = ParseCollectionIndexer(key, TypeCheckingHelper.Instance.GetDictionaryKeyType(dictPropertyInfo.PropertyType));
                child = ((IDictionary)dictPropertyInfo.GetValue(model, null))[indexer];
                return true;
            }
            if (!TryExtractCurrentChildPropertyInfo(model, pathPart, out var propertyInfo))
                return false;
            child = propertyInfo.GetValue(model, null);
            return true;
        }

        private static void ExtractChildObject(object model, string[] pathParts, int pathPartIndex, List<object> result)
        {
            if(pathPartIndex == pathParts.Length)
            {
                result.Add(model);
                return;
            }

            if (!TryExtractCurrentChild(model, pathParts[pathPartIndex], out var currentChild))
                return;
            
            if (currentChild == null)
            {
                result.Add(null);
                return;
            }

            if(TemplateDescriptionHelper.Instance.IsArrayPathPart(pathParts[pathPartIndex]))
            {
                if(!TypeCheckingHelper.Instance.IsEnumerable(currentChild.GetType()))
                    throw new ObjectPropertyExtractionException($"Trying to extract enumerable from non-enumerable property {string.Join(".", pathParts)}");
                foreach(var element in ((IEnumerable)currentChild).Cast<object>())
                    ExtractChildObject(element, pathParts, pathPartIndex + 1, result);
                return;
            }

            ExtractChildObject(currentChild, pathParts, pathPartIndex + 1, result);
        }

        private static object ExtractChildObject(object model, string[] pathParts)
        {
            var result = new List<object>();
            ExtractChildObject(model, pathParts, 0, result);

            if(result.Count == 0)
                throw new ObjectPropertyExtractionException($"Can't find path '{string.Join(".", pathParts)}' in model of type '{model.GetType()}'");

            if(pathParts.Any(TemplateDescriptionHelper.Instance.IsArrayPathPart))
                return result.ToArray();
            return result.Single();
        }

        private static Action<object> ExtractChildModelSetter(object model, string[] pathParts)
        {
            var modelParameter = Expression.Variable(typeof(object));
            var currNodeType = model.GetType();
            Expression currNodeExpression = Expression.Convert(modelParameter, currNodeType);


            var argumentObjectExpression = Expression.Parameter(typeof(object));

            var statements = ExtractChildModelSetterInner(currNodeType, currNodeExpression, argumentObjectExpression, pathParts);
            
            var block = Expression.Block(new ParameterExpression[0], statements);

            var expression = Expression.Lambda<Action<object, object>>(block, modelParameter, argumentObjectExpression);

            var act = expression.Compile();

            return x => act(model, x);
        }

        // todo (mpivko, 30.01.2018): rewrite next 5 methods without expressions
        private static Expression CreateValueInitializationStatement(Expression expression, Type valueType)
        {
            var currNodeConstructor = valueType.GetConstructor(new Type[0]);
            if (currNodeConstructor == null)
                throw new NotSupportedExcelSerializationException($"Object has no parameterless constructor (type - {valueType})");
            var createCurrNodeExpression = Expression.New(currNodeConstructor);
            var assignCurrNodeExpression = Expression.Assign(expression, createCurrNodeExpression);
            return assignCurrNodeExpression;
        }

        private static Expression CreateArrayInitializationStatement(Expression expression, Type arrayItemType, Expression length)
        {
            var createCurrNodeExpression = Expression.NewArrayBounds(arrayItemType, length);
            var assignCurrNodeExpression = Expression.Assign(expression, createCurrNodeExpression);
            return assignCurrNodeExpression;
        }

        private static Expression CreateInitStatement(Expression currNodeIsNullCondition, Expression initializationStatement)
        {
            var maybeInit = Expression.IfThen(currNodeIsNullCondition, initializationStatement);

            var tryBlock = maybeInit;
            var catchBlock = initializationStatement;

            return Expression.TryCatch(Expression.Block(typeof(void), tryBlock), Expression.Catch(typeof(KeyNotFoundException), Expression.Block(typeof(void), catchBlock)));
        }

        private static Expression CreateValueInitStatement(Expression expression, Type valueType)
        {
            var currNodeIsNullCondition = Expression.Equal(expression, Expression.Constant(null, valueType));
            return CreateInitStatement(currNodeIsNullCondition, CreateValueInitializationStatement(expression, valueType));
        }

        private static Expression CreateArrayInitStatement(Expression expression, Type arrayItemType, Expression length)
        {
            var currNodeIsNullCondition = Expression.Equal(expression, Expression.Constant(null, Array.CreateInstance(arrayItemType, 0).GetType()));
            return CreateInitStatement(currNodeIsNullCondition, CreateArrayInitializationStatement(expression, arrayItemType, length));
        }

        [NotNull, ItemNotNull]
        private static List<Expression> ExtractChildModelSetterInner([NotNull] Type currNodeType, [NotNull] Expression currNodeExpression, [NotNull] Expression argumentObjectExpression, [NotNull, ItemNotNull] string[] pathParts)
        {
            var statements = new List<Expression>();

            var needToAssign = true;
            foreach (var (i, part) in pathParts.WithIndices())
            {
                var name = TemplateDescriptionHelper.Instance.GetPathPartName(part);
                
                if(!TypeCheckingHelper.Instance.IsNullable(currNodeType) && i != 0)
                {
                    if(currNodeType.IsArray)
                    {
                        var getLenExpression = Expression.Property(Expression.Convert(argumentObjectExpression, typeof(List<object>)), "Count");
                        statements.Add(CreateArrayInitStatement(currNodeExpression, currNodeType.GetElementType(), getLenExpression));
                    }
                    else
                    {
                        statements.Add(CreateValueInitStatement(currNodeExpression, currNodeType));
                    }
                }

                var newNodeType = currNodeType.GetProperty(name)?.PropertyType;
                currNodeType = newNodeType ?? throw new ObjectPropertyExtractionException($"Type '{currNodeType}' has no property '{name}'");
                currNodeExpression = Expression.Property(currNodeExpression, name);

                if (TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(part))
                {
                    if(TypeCheckingHelper.Instance.IsDictionary(currNodeType))
                    {
                        // todo (mpivko, 24.01.2018): check that it's not a first iteration
                        statements.Add(CreateValueInitStatement(currNodeExpression, currNodeType));

                        var dictKeyType = TypeCheckingHelper.Instance.GetDictionaryKeyType(currNodeType);
                        var indexer = ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), dictKeyType);

                        var dictElementExpression = Expression.Property(currNodeExpression, "Item", Expression.Constant(indexer));
                        var dictValueType = TypeCheckingHelper.Instance.GetDictionaryValueType(currNodeType);

                        currNodeExpression = dictElementExpression;
                        currNodeType = dictValueType; // todo (mpivko, 24.01.2018): it can't be not dict
                    }
                    else if(currNodeType.IsArray)
                    {
                        var arrayItemType = currNodeType.GetElementType() ?? throw new ObjectPropertyExtractionException($"Array of type '{currNodeType}' has no item type");
                        var extendArrayMethod = typeof(ObjectPropertiesExtractor).GetMethod(nameof(ExtendArray), BindingFlags.NonPublic | BindingFlags.Static).MakeGenericMethod(arrayItemType);
                        var indexer = ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), typeof(int));
                        statements.Add(Expression.Assign(currNodeExpression, Expression.Call(extendArrayMethod, currNodeExpression, Expression.Constant((int)indexer + 1))));
                        var arrayItemExpression = Expression.ArrayAccess(currNodeExpression, Expression.Constant(indexer));
                        
                        currNodeExpression = arrayItemExpression;
                        currNodeType = arrayItemType;
                    }
                    else
                    {
                        throw new ObjectPropertyExtractionException("Only dicts and arrays are supported as collections");
                    }
                }
                else if(TemplateDescriptionHelper.Instance.IsArrayPathPart(part))
                {
                    if(currNodeType.IsArray)//todo mpivko potentially it could be imporoved to TypeCheckingHelper.Instance.IsIList(currNodeType))
                    {
                        var itemType = TypeCheckingHelper.Instance.GetEnumerableItemType(currNodeType);
                        
                        var getLenExpression = Expression.Property(Expression.Convert(argumentObjectExpression, typeof(ICollection)), "Count");
                        statements.Add(CreateArrayInitStatement(currNodeExpression, currNodeType.GetElementType(), getLenExpression));

                        var expressionLoopVar = Expression.Variable(typeof(int));

                        var elementExpression = Expression.ArrayAccess(currNodeExpression, expressionLoopVar);
                        var elementInitStatement = CreateValueInitStatement(elementExpression, itemType);
                        var setItemStatements = Expression.Block(ExtractChildModelSetterInner(itemType, elementExpression, GetIndexAccessExpression(argumentObjectExpression, expressionLoopVar), pathParts.Skip(i + 1).ToArray()));
                        
                        var loopBodyExpression = Expression.Block(elementInitStatement, setItemStatements);
                        
                        var loopExpression = ForFromTo(expressionLoopVar, Expression.Constant(0), getLenExpression, loopBodyExpression);

                        statements.Add(loopExpression);

                        needToAssign = false;
                        break;
                    }
                    throw new ObjectPropertyExtractionException("Only array is supported as iterated collection");
                }
            }
            if(needToAssign)
            {
                statements.AddRange(AssignWithTypeChecks(currNodeExpression, currNodeType, argumentObjectExpression));
            }
            return statements;
        }

        private static List<Expression> AssignWithTypeChecks(Expression target, Type targetType, Expression from)
        {
            var res = new List<Expression>();
            var argumentType = Expression.Call(from, typeof(object).GetMethod("GetType"));
            var isFromNull = Expression.Equal(from, Expression.Constant(null));
            var nullToValueViolationCondition = Expression.AndAlso(isFromNull, Expression.IsTrue(Expression.Property(Expression.Constant(targetType), "IsValueType")));
            res.Add(Expression.IfThen(nullToValueViolationCondition, Expression.Throw(Expression.Constant(new ObjectPropertyExtractionException($"Can't assign null to value-type {targetType}")))));
            var wrongArgumentTypeCondition = Expression.AndAlso(Expression.IsFalse(isFromNull), Expression.IsFalse(Expression.Call(Expression.Constant(targetType), typeof(Type).GetMethod("IsAssignableFrom"), argumentType))); // todo (mpivko, 31.01.2018): add test for i
            res.Add(Expression.IfThen(wrongArgumentTypeCondition, Expression.Throw(Expression.Constant(new ObjectPropertyExtractionException($"Can't assign item of type {argumentType} to {targetType}")))));
            var argumentObjectCastExpression = Expression.Convert(from, targetType);
            res.Add(Expression.Assign(target, argumentObjectCastExpression));
            return res;
        }

        private static T[] ExtendArray<T>(T[] array, int length)
        {
            if (array == null)
                return new T[length];
            if(array.Length >= length)
                return array;
            var newArray = new T[length];
            foreach(var (i, item) in array.WithIndices())
                newArray[i] = item;
            return newArray;
        }

        [NotNull]
        private static Expression ForFromTo([NotNull] ParameterExpression loopVar, [NotNull] Expression lInclusive, [NotNull] Expression rExclusive, [NotNull] Expression body)
        {
            var expressionLoopVarInit = Expression.Assign(loopVar, lInclusive);

            var conditionCheckExpression = Expression.GreaterThanOrEqual(loopVar, rExclusive);
            var expressionIncrement = Expression.Assign(loopVar, Expression.Increment(loopVar));

            return Expression.Block(new [] {loopVar}, For(expressionLoopVarInit, conditionCheckExpression, expressionIncrement, body));
        }

        [NotNull]
        private static Expression For([NotNull] Expression initialization, [NotNull] Expression condition, [NotNull] Expression action, [NotNull] Expression body)
        {
            var breakLabel = Expression.Label($"LoopBreak-{Guid.NewGuid()}");

            var conditionCheckExpression = Expression.IfThen(condition, Expression.Goto(breakLabel));
            var loopBodyExpression = Expression.Block(conditionCheckExpression, body, action);
            return Expression.Block(initialization, Expression.Loop(loopBodyExpression, breakLabel));
        }

        [NotNull]
        private static Expression GetIndexAccessExpression([NotNull] Expression target, [NotNull] Expression index)
        {
            var list = Expression.Convert(target, typeof(IList));
            return Expression.Property(list, "Item", index);
        }

        [NotNull]
        private static object ParseCollectionIndexer([NotNull] string collectionIndexer, [NotNull] Type collectionKeyType)
        {
            if(collectionIndexer.StartsWith("\"") && collectionIndexer.EndsWith("\""))
            {
                if (collectionKeyType != typeof(string))
                    throw new ObjectPropertyExtractionException($"Collection with '{collectionKeyType}' keys was indexed by {typeof(string)}");
                return collectionIndexer.Substring(1, collectionIndexer.Length - 2);
            }
            if(int.TryParse(collectionIndexer, out var intIndexer))
            {
                if (collectionKeyType != typeof(int))
                    throw new ObjectPropertyExtractionException($"Collection with '{collectionKeyType}' keys was indexed by {typeof(int)}");
                return intIndexer;
            }
            throw new ObjectPropertyExtractionException("Only strings and ints are supported as collection indexers");
        }
    }
}