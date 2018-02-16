using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.Linq;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public static class ObjectPropertySettersExtractor
    {
        [NotNull]
        public static Action<object> ExtractChildObjectSetter([NotNull] object model, [NotNull] ExcelTemplatePath path)
        {
            var action = childObjectSettersCache.GetOrAdd((model.GetType(), path), x => ExtractChildModelSetter(x.Item1, x.Item2.PartsWithIndexers));
            return x => action(model, x);
        }

        [NotNull]
        private static Action<object, object> ExtractChildModelSetter([NotNull] Type modelType, [NotNull, ItemNotNull] string[] pathParts)
        {
            var targetModel = Expression.Variable(typeof(object));
            var valueToSet = Expression.Parameter(typeof(object));
            var currentModelNode = Expression.Convert(targetModel, modelType);

            var statements = BuildExtractionOfChildModelSetter(modelType, currentModelNode, valueToSet, pathParts);
            var block = Expression.Block(new ParameterExpression[0], statements);
            var expression = Expression.Lambda<Action<object, object>>(block, targetModel, valueToSet);

            return expression.Compile();
        }

        [NotNull, ItemNotNull]
        private static List<Expression> BuildExtractionOfChildModelSetter([NotNull] Type currNodeType, [NotNull] Expression currNodeExpression, [NotNull] Expression valueToSetExpression, [NotNull, ItemNotNull] string[] pathParts)
        {
            var statements = new List<Expression>();

            foreach(var (i, part) in pathParts.WithIndices())
            {
                var name = TemplateDescriptionHelper.Instance.GetPathPartName(part);

                var newNodeType = currNodeType.GetProperty(name)?.PropertyType;
                currNodeType = newNodeType ?? throw new ObjectPropertyExtractionException($"Type '{currNodeType}' has no property '{name}'");
                currNodeExpression = Expression.Property(currNodeExpression, name);

                if(TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(part))
                {
                    List<Expression> statementsToAdd;
                    (currNodeExpression, currNodeType, statementsToAdd) = BuildExpandingOfCollectionAccessPart(currNodeExpression, currNodeType, part);
                    statements.AddRange(statementsToAdd);
                }
                else if(TemplateDescriptionHelper.Instance.IsArrayPathPart(part))
                {
                    var statementsToAdd = BuildExpandingOfArrayPart(currNodeExpression, currNodeType, valueToSetExpression, pathParts.Skip(i + 1).ToArray());
                    statements.AddRange(statementsToAdd);
                    return statements;
                }
                else if(!TypeCheckingHelper.Instance.IsNullable(currNodeType) && i != pathParts.Length - 1)
                {
                    if(currNodeType.IsArray)
                    {
                        var getLenExpression = Expression.Property(Expression.Convert(valueToSetExpression, typeof(List<object>)), "Count");
                        statements.Add(ExpressionPrimitives.CreateArrayInitStatement(currNodeExpression, currNodeType.GetElementType(), getLenExpression));
                    }
                    else
                    {
                        statements.Add(ExpressionPrimitives.CreateValueInitStatement(currNodeExpression, currNodeType));
                    }
                }
            }
            statements.Add(ExpressionPrimitives.AssignWithTypeCheckings(currNodeExpression, currNodeType, valueToSetExpression));
            return statements;
        }
        
        private static (Expression currNodeExpression, Type currNodeType, List<Expression> statements) BuildExpandingOfCollectionAccessPart([NotNull] Expression currNodeExpression, [NotNull] Type currNodeType, [NotNull] string part)
        {
            var statements = new List<Expression>();
            if(TypeCheckingHelper.Instance.IsDictionary(currNodeType))
            {
                statements.Add(ExpressionPrimitives.CreateValueInitStatement(currNodeExpression, currNodeType));

                var dictKeyType = TypeCheckingHelper.Instance.GetDictionaryKeyType(currNodeType);
                var indexer = TemplateDescriptionHelper.ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), dictKeyType);

                var dictElementExpression = Expression.Property(currNodeExpression, "Item", Expression.Constant(indexer));
                var dictValueType = TypeCheckingHelper.Instance.GetDictionaryValueType(currNodeType);

                statements.Add(ExpressionPrimitives.CreateDictValueInitStatement(currNodeExpression, dictKeyType, dictValueType, indexer));

                currNodeExpression = dictElementExpression;
                currNodeType = dictValueType;
            }
            else if(currNodeType.IsArray)
            {
                var arrayItemType = currNodeType.GetElementType() ?? throw new ObjectPropertyExtractionException($"Array of type '{currNodeType}' has no item type");
                var indexer = TemplateDescriptionHelper.ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), typeof(int));
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
            var statements = new List<Expression>();
            if(currNodeType.IsArray)
            {
                var itemType = TypeCheckingHelper.Instance.GetEnumerableItemType(currNodeType);

                var getLenExpression = Expression.Property(Expression.Convert(valueToSetExpression, typeof(ICollection)), "Count");
                statements.Add(ExpressionPrimitives.CreateArrayInitStatement(currNodeExpression, itemType, getLenExpression));

                var expressionLoopVar = Expression.Variable(typeof(int));

                var elementExpression = Expression.ArrayAccess(currNodeExpression, expressionLoopVar);
                var elementInitStatement = ExpressionPrimitives.CreateValueInitStatement(elementExpression, itemType);
                var setItemStatements = Expression.Block(BuildExtractionOfChildModelSetter(itemType, elementExpression, ExpressionPrimitives.GetIndexAccessExpression(valueToSetExpression, expressionLoopVar), pathParts));

                var loopBodyExpression = Expression.Block(elementInitStatement, setItemStatements);

                var loopExpression = ExpressionPrimitives.ForFromTo(expressionLoopVar, Expression.Constant(0), getLenExpression, loopBodyExpression);

                statements.Add(loopExpression);
                return statements;
            }
            throw new ObjectPropertyExtractionException("Only array is supported as iterated collection");
        }

        [NotNull, ItemNotNull]
        private static readonly ConcurrentDictionary<(Type, ExcelTemplatePath), Action<object, object>> childObjectSettersCache = new ConcurrentDictionary<(Type, ExcelTemplatePath), Action<object, object>>();
    }
}