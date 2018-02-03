using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.Linq;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public sealed class ObjectPropertiesExtractor
    {
        private ObjectPropertiesExtractor()
        {
        }

        [CanBeNull]
        public object ExtractChildObject(object model, ExcelTemplateExpression expression)
        {
            return ExtractChildObject(model, expression.ChildObjectPath.PartsWithIndexers);
        }

        [CanBeNull]
        public object ExtractChildObjectViaPath(object model, ExcelTemplatePath path)
        {
            return ExtractChildObject(model, path.PartsWithIndexers);
        }

        [NotNull]
        public static Type ExtractChildObjectTypeFromPath([NotNull] Type modelType, [NotNull] ExcelTemplatePath path)
        {
            var currType = modelType;

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

        public static (string, string) SplitForEnumerableExpansion([NotNull] ExcelTemplateExpression expression)
        {
            if (!expression.ChildObjectPath.HasArrayAccess)
                throw new ObjectPropertyExtractionException($"Expression needs enumerable expansion but has no part with '[]' (expression - '{expression}')");
            var parts = expression.ChildObjectPath.PartsWithIndexers;
            var firstPartLen = parts.TakeWhile(x => !x.EndsWith("[]")).Count() + 1;
            return (string.Join(".", parts.Take(firstPartLen)), string.Join(".", parts.Skip(firstPartLen)));
        }
        
        [NotNull]
        public static Action<object> ExtractChildObjectSetter([NotNull] object model, [NotNull] ExcelTemplateExpression expression)
        {
            return ExtractChildModelSetter(model, expression.ChildObjectPath.PartsWithIndexers);
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
                if (!TryExtractCurrentChildPropertyInfo(model, name, out var collectionPropertyInfo))
                    return false;
                if(TypeCheckingHelper.Instance.IsDictionary(collectionPropertyInfo.PropertyType))
                {
                    var indexer = ParseCollectionIndexer(key, TypeCheckingHelper.Instance.GetDictionaryKeyType(collectionPropertyInfo.PropertyType));
                    var dict = collectionPropertyInfo.GetValue(model, null);
                    if (dict == null)
                        return true;
                    child = ((IDictionary)dict)[indexer];
                    return true;
                }
                if(TypeCheckingHelper.Instance.IsIList(collectionPropertyInfo.PropertyType))
                {
                    var indexer = (int)ParseCollectionIndexer(key, typeof(int));
                    var list = collectionPropertyInfo.GetValue(model, null);
                    if (list == null)
                        return true;
                    child = ((IList)list)[indexer];
                    return true;
                }
                throw new ObjectPropertyExtractionException($"Unexpected child type: expected dictionary or array (pathPath='{pathPart}'), but model is '{collectionPropertyInfo.PropertyType}' in '{model.GetType()}'");
                
            }
            if (!TryExtractCurrentChildPropertyInfo(model, pathPart, out var propertyInfo))
                return false;
            child = propertyInfo.GetValue(model, null);
            return true;
        }

        private static bool TryExtractChildObject(object model, string[] pathParts, int pathPartIndex, out object result)
        {
            if(pathPartIndex == pathParts.Length)
            {
                result = model;
                return true;
            }

            if(!TryExtractCurrentChild(model, pathParts[pathPartIndex], out var currentChild))
            {
                result = null;
                return false;
            }

            if (currentChild == null)
            {
                result = null;
                return true;
            }

            if(TemplateDescriptionHelper.Instance.IsArrayPathPart(pathParts[pathPartIndex]))
            {
                if(!TypeCheckingHelper.Instance.IsEnumerable(currentChild.GetType()))
                    throw new ObjectPropertyExtractionException($"Trying to extract enumerable from non-enumerable property {string.Join(".", pathParts)}");
                var resultList = new List<object>();
                foreach(var element in ((IEnumerable)currentChild).Cast<object>())
                {
                    if (element == null)
                        resultList.Add(null);
                    else
                    {
                        if (!TryExtractChildObject(element, pathParts, pathPartIndex + 1, out var resultItem))
                        {
                            result = null;
                            return false;
                        }
                        resultList.Add(resultItem);
                    }
                }
                result = resultList.ToArray();
                return true;
            }

            return TryExtractChildObject(currentChild, pathParts, pathPartIndex + 1, out result);
        }

        [CanBeNull]
        private static object ExtractChildObject(object model, string[] pathParts)
        {
            if (!TryExtractChildObject(model, pathParts, 0, out var result))
                throw new ObjectPropertyExtractionException($"Can't find path '{string.Join(".", pathParts)}' in model of type '{model.GetType()}'");
            
            return result;
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
        
        private static T InitValue<T>(T currentValue)
            where T : new()
        {
            if (typeof(T).IsValueType)
                return currentValue;
            if (currentValue != null)
                return currentValue;
            return new T();
        }

        private static T[] InitArray<T>(T[] currentValue, int length)
            where T : new()
        {
            if (currentValue != null)
                return currentValue;
            return new T[length];
        }

        private static void InitDict<TKey, TValue>(Dictionary<TKey, TValue> dict, object indexer, Func<TValue> createValue)
        {
            if (!(indexer is TKey realIndexer))
                throw new ObjectPropertyExtractionException($"Indexer type '{indexer?.GetType().ToString() ?? "NULL"}' does not match dictionary key type '{typeof(TKey)}'");
            if (indexer == null)
                throw new ObjectPropertyExtractionException("Can't use null as dict key");
            if (!dict.ContainsKey(realIndexer))
                dict[realIndexer] = createValue();
        }

        private static void InitPrimitiveDict<TKey, TValue>(Dictionary<TKey, TValue> dict, object indexer)
        {
            InitDict(dict, indexer, () => default);
        }

        private static void InitClassDict<TKey, TValue>(Dictionary<TKey, TValue> dict, object indexer)
            where TValue : new()
        {
            InitDict(dict, indexer, () => new TValue());
        }

        private static Expression CreateValueInitStatement(Expression expression, Type valueType)
        {
            var method = GetGenericMethod(nameof(InitValue), valueType);
            return Expression.Assign(expression, Expression.Call(method, expression));
        }

        private static Expression CreateArrayInitStatement(Expression expression, Type arrayItemType, Expression length)
        {
            var method = GetGenericMethod(nameof(InitArray), arrayItemType);
            return Expression.Assign(expression, Expression.Call(method, expression, length));
        }

        private static Expression CreateDictValueInitStatement(Expression dict, Type dictKeyType, Type dictValueType, object indexer)
        {
            MethodInfo method;
            if (dictValueType.IsValueType || dictValueType == typeof(string))
                method = GetGenericMethod(nameof(InitPrimitiveDict), dictKeyType, dictValueType);
            else
                method = GetGenericMethod(nameof(InitClassDict), dictKeyType, dictValueType);
            return Expression.Call(method, dict, Expression.Convert(Expression.Constant(indexer), typeof(object)));
        }
        
        [NotNull, ItemNotNull]
        private static List<Expression> ExtractChildModelSetterInner([NotNull] Type currNodeType, [NotNull] Expression currNodeExpression, [NotNull] Expression argumentObjectExpression, [NotNull, ItemNotNull] string[] pathParts)
        {
            var statements = new List<Expression>();
            
            foreach (var (i, part) in pathParts.WithIndices())
            {
                var name = TemplateDescriptionHelper.Instance.GetPathPartName(part);
                
                var newNodeType = currNodeType.GetProperty(name)?.PropertyType;
                currNodeType = newNodeType ?? throw new ObjectPropertyExtractionException($"Type '{currNodeType}' has no property '{name}'");
                currNodeExpression = Expression.Property(currNodeExpression, name);

                if (TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(part))
                {
                    List<Expression> statementsToAdd;
                    (currNodeExpression, currNodeType, statementsToAdd) = ParseCollectionAccessPart(currNodeExpression, currNodeType, part);
                    statements.AddRange(statementsToAdd);
                }
                else if(TemplateDescriptionHelper.Instance.IsArrayPathPart(part))
                {
                    var statementsToAdd = ParseArrayPart(currNodeExpression, currNodeType, argumentObjectExpression, pathParts.Skip(i + 1).ToArray());
                    statements.AddRange(statementsToAdd);
                    return statements;
                }
                else
                {
                    if (!TypeCheckingHelper.Instance.IsNullable(currNodeType) && i != pathParts.Length - 1)
                    {
                        if (currNodeType.IsArray)
                        {
                            var getLenExpression = Expression.Property(Expression.Convert(argumentObjectExpression, typeof(List<object>)), "Count");
                            statements.Add(CreateArrayInitStatement(currNodeExpression, currNodeType.GetElementType(), getLenExpression));
                        }
                        else
                        {
                            statements.Add(CreateValueInitStatement(currNodeExpression, currNodeType));
                        }
                    }
                }
            }
            statements.Add(AssignWithTypeChecks(currNodeExpression, currNodeType, argumentObjectExpression));
            return statements;
        }

        private static (Expression currNodeExpression, Type currNodeType, List<Expression> statements) ParseCollectionAccessPart(Expression currNodeExpression, Type currNodeType, string part)
        {
            var statements = new List<Expression>();
            if (TypeCheckingHelper.Instance.IsDictionary(currNodeType))
            {
                statements.Add(CreateValueInitStatement(currNodeExpression, currNodeType));

                var dictKeyType = TypeCheckingHelper.Instance.GetDictionaryKeyType(currNodeType);
                var indexer = ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), dictKeyType);

                var dictElementExpression = Expression.Property(currNodeExpression, "Item", Expression.Constant(indexer));
                var dictValueType = TypeCheckingHelper.Instance.GetDictionaryValueType(currNodeType);

                statements.Add(CreateDictValueInitStatement(currNodeExpression, dictKeyType, dictValueType, indexer));

                currNodeExpression = dictElementExpression;
                currNodeType = dictValueType;
            }
            else if (currNodeType.IsArray)
            {
                var arrayItemType = currNodeType.GetElementType() ?? throw new ObjectPropertyExtractionException($"Array of type '{currNodeType}' has no item type");
                var extendArrayMethod = GetGenericMethod(nameof(ExtendArray), arrayItemType);
                var indexer = ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), typeof(int));
                statements.Add(Expression.Assign(currNodeExpression, Expression.Call(extendArrayMethod, currNodeExpression, Expression.Constant((int)indexer + 1))));
                var arrayItemExpression = Expression.ArrayAccess(currNodeExpression, Expression.Constant(indexer));

                currNodeExpression = arrayItemExpression;
                currNodeType = arrayItemType;

                statements.Add(CreateValueInitStatement(currNodeExpression, currNodeType));
            }
            else
            {
                throw new ObjectPropertyExtractionException("Only dicts and arrays are supported as collections");
            }
            return (currNodeExpression, currNodeType, statements);
        }

        private static List<Expression> ParseArrayPart(Expression currNodeExpression, Type currNodeType, Expression argumentObjectExpression, string[] pathParts)
        {
            var statements = new List<Expression>();
            if (currNodeType.IsArray) //todo mpivko potentially it could be imporoved to TypeCheckingHelper.Instance.IsIList(currNodeType))
            {
                var itemType = TypeCheckingHelper.Instance.GetEnumerableItemType(currNodeType);

                var getLenExpression = Expression.Property(Expression.Convert(argumentObjectExpression, typeof(ICollection)), "Count");
                statements.Add(CreateArrayInitStatement(currNodeExpression, itemType, getLenExpression));

                var expressionLoopVar = Expression.Variable(typeof(int));

                var elementExpression = Expression.ArrayAccess(currNodeExpression, expressionLoopVar);
                var elementInitStatement = CreateValueInitStatement(elementExpression, itemType);
                var setItemStatements = Expression.Block(ExtractChildModelSetterInner(itemType, elementExpression, GetIndexAccessExpression(argumentObjectExpression, expressionLoopVar), pathParts));

                var loopBodyExpression = Expression.Block(elementInitStatement, setItemStatements);

                var loopExpression = ForFromTo(expressionLoopVar, Expression.Constant(0), getLenExpression, loopBodyExpression);

                statements.Add(loopExpression);
                return statements;
            }
            throw new ObjectPropertyExtractionException("Only array is supported as iterated collection");
        }

        private static Expression AssignWithTypeChecks(Expression target, Type targetType, Expression from)
        {
            var smartCastMethod = GetGenericMethod(nameof(SmartCast), targetType);
            return Expression.Assign(target, Expression.Call(smartCastMethod, @from));
        }

        private static T SmartCast<T>(object from)
        {
            if(from is T res)
                return res;
            if(from == null && !typeof(T).IsValueType)
                return default;
            throw new ObjectPropertyExtractionException($"Can't assign item of type '{from?.GetType().ToString() ?? "NULL"}' to target of type '{typeof(T)}'");
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

        private static MethodInfo GetMethod(string name)
        {
            var method = typeof(ObjectPropertiesExtractor).GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static);
            if (method == null)
                throw new InvalidProgramStateException($"Method '{name}' not found in '{typeof(ObjectPropertiesExtractor)}'");
            return method;
        }

        private static MethodInfo GetGenericMethod(string name, params Type[] genericTypes)
        {
            return GetMethod(name).MakeGenericMethod(genericTypes);
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