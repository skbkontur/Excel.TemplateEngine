using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq.Expressions;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers
{
    internal static class ObjectConversionGenerator
    {
        [NotNull]
        public static Func<Dictionary<ExcelTemplatePath, object>, object> BuildDictToObject([NotNull] ExcelTemplatePath[] objectProps, [NotNull] Type objectType)
        {
            var dictType = typeof(Dictionary<ExcelTemplatePath, object>);
            var objectDict = Expression.Parameter(dictType);

            var newObject = Expression.Variable(typeof(object));
            var typedObject = Expression.Convert(newObject, objectType);

            var objectConstructor = objectType.GetConstructor(Array.Empty<Type>());
            var initObject = Expression.Assign(newObject, Expression.New(objectConstructor!));

            var minExpressionCount = objectProps.Length + 2;
            var expressions = new List<Expression>(minExpressionCount) {initObject};

            var dictIndexer = dictType.GetProperty("Item");
            foreach (var prop in objectProps)
            {
                var setProp = ChildSetterGenerator.BuildChildSetter(objectType, typedObject, Expression.MakeIndex(objectDict, dictIndexer, new[] {Expression.Constant(prop)}), prop.PartsWithIndexers);
                expressions.AddRange(setProp);
            }

            expressions.Add(newObject);

            var block = Expression.Block(new[] {newObject}, expressions);
            return Expression.Lambda<Func<Dictionary<ExcelTemplatePath, object>, object>>(block, objectDict)
                             .Compile();
        }

        [NotNull]
        private static readonly ConcurrentDictionary<Type, Func<Dictionary<ExcelTemplatePath, object>, object>> dictToObjectCache = new ConcurrentDictionary<Type, Func<Dictionary<ExcelTemplatePath, object>, object>>();
    }
}