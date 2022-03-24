using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.Exceptions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers
{
    internal static class ObjectPropertySettersExtractor
    {
        [NotNull]
        public static Action<object> ExtractChildObjectSetter([NotNull] object model, [NotNull] ExcelTemplatePath path)
        {
            var action = childObjectSettersCache.GetOrAdd((model.GetType(), path), x => ExtractChildModelSetter(x.type, x.path.PartsWithIndexers));
            return x => action(model, x);
        }

        /// <summary>
        ///     Generates function to add item to list by given path.
        ///     Generated function accepts new item in Dictionary<ExcelTemplatePath, object> format, where Kay is item property path and Value is object to set this property to.
        /// </summary>
        /// <param name="model">Base object.</param>
        /// <param name="pathToList"></param>
        /// <param name="relativeItemProps">List of paths to item properties to set.</param>
        /// <returns></returns>
        [NotNull]
        public static Action<Dictionary<ExcelTemplatePath, object>> GenerateChildListItemAdder([NotNull] object model, [NotNull] ExcelTemplatePath pathToList, [NotNull] ExcelTemplatePath[] relativeItemProps)
        {
            var modelType = model.GetType();
            var itemDictType = typeof(Dictionary<ExcelTemplatePath, object>);

            var modelParam = Expression.Parameter(typeof(object));
            var newItemDict = Expression.Parameter(itemDictType);
            var typedModelParam = Expression.Convert(modelParam, modelType);

            var clearPathToList = pathToList.WithoutArrayAccess();
            var listType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, clearPathToList);
            var listConstructor = listType.GetConstructor(Array.Empty<Type>());
            // valueToSetExpression param is never set
            var initListStatements = BuildChildSetter(modelType, typedModelParam, Expression.New(listConstructor!), pathToList.PartsWithIndexers);

            var itemType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, pathToList);
            var itemConstructor = itemType.GetConstructor(Array.Empty<Type>());

            var newItem = Expression.Variable(typeof(object));
            var typedItem = Expression.Convert(newItem, itemType);
            var initNewItem = Expression.Assign(newItem, Expression.New(itemConstructor!));

            var dictIndexer = itemDictType.GetProperty("Item");
            var minListLength = relativeItemProps.Select(x => x.PartsWithIndexers.Length)
                                                 .Sum();
            var setItemProps = new List<Expression>(minListLength + 1);
            setItemProps.Add(initNewItem);
            foreach (var prop in relativeItemProps)
            {
                var setProp = BuildChildSetter(itemType, typedItem, Expression.MakeIndex(newItemDict, dictIndexer, new[] {Expression.Constant(prop)}), prop.PartsWithIndexers);
                setItemProps.AddRange(setProp);
            }

            var addItemToList = GenerateAddItemToList(typedModelParam, listType, clearPathToList.PartsWithIndexers, typedItem);

            var block = Expression.Block(new[] {newItem}, initListStatements.Concat(setItemProps).Concat(new[] {addItemToList}));
            var wholeAction = Expression.Lambda<Action<object, Dictionary<ExcelTemplatePath, object>>>(block, modelParam, newItemDict)
                                        .Compile();
            return x => wholeAction(model, x);
        }

        private static MethodCallExpression GenerateAddItemToList(Expression targetModel, Type listType, string[] clearPathToList, Expression newItem)
        {
            var listExpression = Expression.Property(targetModel, clearPathToList.First());
            listExpression = clearPathToList.Skip(1).Aggregate(listExpression, Expression.Property);

            var addMethodInfo = listType.GetMethod("Add");
            return Expression.Call(listExpression, addMethodInfo, newItem);
        }

        [NotNull]
        private static Action<object, object> ExtractChildModelSetter([NotNull] Type modelType, [NotNull, ItemNotNull] string[] pathParts)
        {
            var targetModel = Expression.Variable(typeof(object));
            var valueToSet = Expression.Parameter(typeof(object));
            var currentModelNode = Expression.Convert(targetModel, modelType);

            var statements = BuildChildSetter(modelType, currentModelNode, valueToSet, pathParts);
            var block = Expression.Block(new ParameterExpression[0], statements);
            var expression = Expression.Lambda<Action<object, object>>(block, targetModel, valueToSet);

            return expression.Compile();
        }

        [NotNull, ItemNotNull]
        private static List<Expression> BuildChildSetter([NotNull] Type currNodeType, [NotNull] Expression currNodeExpression, [NotNull] Expression valueToSetExpression, [NotNull, ItemNotNull] string[] pathParts)
        {
            var statements = new List<Expression>(pathParts.Length);

            for (var partIndex = 0; partIndex < pathParts.Length; ++partIndex)
            {
                var name = TemplateDescriptionHelper.GetPathPartName(pathParts[partIndex]);

                var newNodeType = currNodeType.GetProperty(name)?.PropertyType;
                currNodeType = newNodeType ?? throw new ObjectPropertyExtractionException($"Type '{currNodeType}' has no property '{name}'");
                currNodeExpression = Expression.Property(currNodeExpression, name);

                if (TemplateDescriptionHelper.IsCollectionAccessPathPart(pathParts[partIndex]))
                {
                    List<Expression> statementsToAdd;
                    (currNodeExpression, currNodeType, statementsToAdd) = BuildExpandingOfCollectionAccessPart(currNodeExpression, currNodeType, pathParts[partIndex]);
                    statements.AddRange(statementsToAdd);
                }
                else if (TemplateDescriptionHelper.IsArrayPathPart(pathParts[partIndex]))
                {
                    if (currNodeType.IsArray)
                    {
                        var statementsToAdd = BuildExpandingOfArrayPart(currNodeExpression, currNodeType, valueToSetExpression, pathParts.Skip(partIndex + 1).ToArray());
                        statements.AddRange(statementsToAdd);
                    }
                    else if (typeof(IList).IsAssignableFrom(currNodeType))
                    {
                        var itemType = TypeCheckingHelper.GetEnumerableItemType(currNodeType);
                        statements.Add(ExpressionPrimitives.CreateListInitStatement(currNodeExpression, itemType));
                    }
                    else
                        throw new ObjectPropertyExtractionException("Only array and list is supported as iterated collections");
                    return statements;
                }
                else if (!TypeCheckingHelper.IsNullable(currNodeType) && partIndex != pathParts.Length - 1)
                {
                    statements.Add(ExpressionPrimitives.CreateValueInitStatement(currNodeExpression, currNodeType));
                }
            }
            statements.Add(ExpressionPrimitives.AssignWithTypeCheckings(currNodeExpression, currNodeType, valueToSetExpression));
            return statements;
        }

        private static (Expression currNodeExpression, Type currNodeType, List<Expression> statements) BuildExpandingOfCollectionAccessPart([NotNull] Expression currNodeExpression, [NotNull] Type currNodeType, [NotNull] string part)
        {
            var statements = new List<Expression>();
            if (TypeCheckingHelper.IsDictionary(currNodeType))
            {
                statements.Add(ExpressionPrimitives.CreateValueInitStatement(currNodeExpression, currNodeType));

                var (dictKeyType, dictValueType) = TypeCheckingHelper.GetDictionaryGenericTypeArguments(currNodeType);
                var indexer = TemplateDescriptionHelper.ParseCollectionIndexerOrThrow(TemplateDescriptionHelper.GetCollectionAccessPathPartIndex(part), dictKeyType);

                var dictElementExpression = Expression.Property(currNodeExpression, "Item", Expression.Constant(indexer));

                statements.Add(ExpressionPrimitives.CreateDictValueInitStatement(currNodeExpression, dictKeyType, dictValueType, indexer));

                currNodeExpression = dictElementExpression;
                currNodeType = dictValueType;
            }
            else if (currNodeType.IsArray)
            {
                var arrayItemType = currNodeType.GetElementType() ?? throw new ObjectPropertyExtractionException($"Array of type '{currNodeType}' has no item type");
                var indexer = TemplateDescriptionHelper.ParseCollectionIndexerOrThrow(TemplateDescriptionHelper.GetCollectionAccessPathPartIndex(part), typeof(int));
                statements.Add(ExpressionPrimitives.CreateArrayExtendStatement(currNodeExpression, Expression.Constant((int)indexer + 1), arrayItemType));
                var arrayItemExpression = Expression.ArrayAccess(currNodeExpression, Expression.Constant(indexer));

                currNodeExpression = arrayItemExpression;
                currNodeType = arrayItemType;

                statements.Add(ExpressionPrimitives.CreateValueInitStatement(currNodeExpression, currNodeType));
            }
            else
            {
                throw new ObjectPropertyExtractionException("Only dicts and arrays are supported as collections");
            }
            return (currNodeExpression, currNodeType, statements);
        }

        [NotNull, ItemNotNull]
        private static List<Expression> BuildExpandingOfArrayPart([NotNull] Expression currNodeExpression, [NotNull] Type currNodeType, [NotNull] Expression valueToSetExpression, [NotNull, ItemNotNull] string[] pathParts)
        {
            var statements = new List<Expression>(2);
            var itemType = TypeCheckingHelper.GetEnumerableItemType(currNodeType);

            var getLenExpression = Expression.Property(Expression.Convert(valueToSetExpression, typeof(ICollection)), nameof(ICollection.Count));
            statements.Add(ExpressionPrimitives.CreateArrayInitStatement(currNodeExpression, itemType, getLenExpression));
            var expressionLoopVar = Expression.Variable(typeof(int));

            var elementExpression = Expression.ArrayAccess(currNodeExpression, expressionLoopVar);
            var elementInitStatement = ExpressionPrimitives.CreateValueInitStatement(elementExpression, itemType);
            var setItemStatements = Expression.Block(BuildChildSetter(itemType, elementExpression, ExpressionPrimitives.GetIndexAccessExpression(valueToSetExpression, expressionLoopVar), pathParts));

            var loopBodyExpression = Expression.Block(elementInitStatement, setItemStatements);

            var loopExpression = ExpressionPrimitives.ForFromTo(expressionLoopVar, Expression.Constant(0), getLenExpression, loopBodyExpression);

            statements.Add(loopExpression);
            return statements;
        }

        [NotNull]
        private static readonly ConcurrentDictionary<(Type type, ExcelTemplatePath path), Action<object, object>> childObjectSettersCache = new ConcurrentDictionary<(Type, ExcelTemplatePath), Action<object, object>>();
    }
}