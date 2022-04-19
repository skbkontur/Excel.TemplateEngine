using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;

using JetBrains.Annotations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers
{
    internal static class ObjectChildSetterFactory
    {
        [NotNull]
        public static Action<object, object> GetChildObjectSetter([NotNull] Type modelType, [NotNull] ExcelTemplatePath path)
            => childObjectSettersCache.GetOrAdd((modelType, path), x => ChildSetterGenerator.BuildChildSetterLambda(x.type, x.path).Compile());

        [NotNull]
        public static Action<object, object> GetEnumerableSetter([NotNull] Type modelType,
                                                                 [NotNull] ExcelTemplatePath pathToEnumerable,
                                                                 [NotNull] Type enumerableType,
                                                                 [NotNull] Type enumerableItemType)
            => childObjectSettersCache.GetOrAdd((modelType, pathToEnumerable), x => BuildEnumerableSetter(x.type, x.path, enumerableType, enumerableItemType));

        private static Action<object, object> BuildEnumerableSetter([NotNull] Type modelType,
                                                                    [NotNull] ExcelTemplatePath pathToEnumerable,
                                                                    [NotNull] Type enumerableType,
                                                                    [NotNull] Type listItemType)
        {
            var setter = ChildSetterGenerator.BuildChildSetterLambda(modelType, pathToEnumerable);

            var parent = Expression.Parameter(typeof(object));
            var child = Expression.Parameter(typeof(object));
            var typedChild = Expression.Convert(child, typeof(IEnumerable<object>));

            MethodInfo castList;
            if (enumerableType.IsArray)
                castList = ExpressionPrimitives.GetGenericMethod(typeof(ObjectChildSetterFactory), nameof(CastToArray), listItemType);
            else if (TypeCheckingHelper.IsList(enumerableType))
                castList = ExpressionPrimitives.GetGenericMethod(typeof(ObjectChildSetterFactory), nameof(CastToList), listItemType);
            else
                throw new ArgumentException("Only Array and List is supported.");

            var castCall = Expression.Call(castList, typedChild);
            var setterInvocation = Expression.Invoke(setter, parent, castCall);

            var block = Expression.Block(Array.Empty<ParameterExpression>(), setterInvocation);
            return Expression.Lambda<Action<object, object>>(block, parent, child)
                             .Compile();
        }

        [NotNull]
        private static T[] CastToArray<T>([NotNull] IEnumerable<object> list)
        {
            return list.Cast<T>().ToArray();
        }

        [NotNull]
        private static List<T> CastToList<T>([NotNull] IEnumerable<object> list)
        {
            return list.Cast<T>().ToList();
        }

        [NotNull]
        private static readonly ConcurrentDictionary<(Type type, ExcelTemplatePath path), Action<object, object>> childObjectSettersCache = new ConcurrentDictionary<(Type, ExcelTemplatePath), Action<object, object>>();
    }
}