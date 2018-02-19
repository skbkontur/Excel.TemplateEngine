using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq.Expressions;
using System.Reflection;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public static class ExpressionPrimitives
    {
        [NotNull]
        public static Expression ForFromTo([NotNull] ParameterExpression loopVar, [NotNull] Expression lInclusive, [NotNull] Expression rExclusive, [NotNull] Expression body)
        {
            var expressionLoopVarInit = Expression.Assign(loopVar, lInclusive);

            var conditionCheckExpression = Expression.GreaterThanOrEqual(loopVar, rExclusive);
            var expressionIncrement = Expression.Assign(loopVar, Expression.Increment(loopVar));

            return Expression.Block(new[] {loopVar}, For(expressionLoopVarInit, conditionCheckExpression, expressionIncrement, body));
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
        public static Expression GetIndexAccessExpression([NotNull] Expression target, [NotNull] Expression index)
        {
            var list = Expression.Convert(target, typeof(IList));
            return Expression.Property(list, "Item", index);
        }

        [NotNull]
        public static Expression CreateValueInitStatement([NotNull] Expression expression, [NotNull] Type valueType)
        {
            var method = GetGenericMethod(typeof(ExpressionPrimitives), nameof(InitValue), valueType);
            return Expression.Assign(expression, Expression.Call(method, expression));
        }

        [NotNull]
        public static Expression CreateArrayInitStatement([NotNull] Expression expression, [NotNull] Type arrayItemType, [NotNull] Expression length)
        {
            var method = GetGenericMethod(typeof(ExpressionPrimitives), nameof(InitArray), arrayItemType);
            return Expression.Assign(expression, Expression.Call(method, expression, length));
        }

        [NotNull]
        public static Expression CreateDictValueInitStatement([NotNull] Expression dict, [NotNull] Type dictKeyType, [NotNull] Type dictValueType, [NotNull] object indexer)
        {
            MethodInfo method;
            if(dictValueType.IsValueType || dictValueType == typeof(string))
                method = GetGenericMethod(typeof(ExpressionPrimitives), nameof(InitPrimitiveDict), dictKeyType, dictValueType);
            else
                method = GetGenericMethod(typeof(ExpressionPrimitives), nameof(InitClassDict), dictKeyType, dictValueType);
            return Expression.Call(method, dict, Expression.Convert(Expression.Constant(indexer), typeof(object)));
        }

        [NotNull]
        public static Expression CreateArrayExtendStatement([NotNull] Expression array, [NotNull] Expression desiredLength, [NotNull] Type arrayItemType)
        {
            var extendArrayMethod = GetGenericMethod(typeof(ExpressionPrimitives), nameof(ExtendArray), arrayItemType);
            return Expression.Assign(array, Expression.Call(extendArrayMethod, array, desiredLength));
        }

        [NotNull]
        public static Expression AssignWithTypeCheckings([NotNull] Expression target, [NotNull] Type targetType, [NotNull] Expression from)
        {
            var castWithTypeCheckingsMethod = GetGenericMethod(typeof(ExpressionPrimitives), nameof(CastWithTypeCheckings), targetType);
            return Expression.Assign(target, Expression.Call(castWithTypeCheckingsMethod, from));
        }

        [NotNull]
        private static T[] ExtendArray<T>([CanBeNull] T[] array, int length)
        {
            if(array == null)
                return new T[length];
            if(array.Length >= length)
                return array;
            var newArray = new T[length];
            for(var i = 0; i < array.Length; i++)
                newArray[i] = array[i];
            return newArray;
        }

        [CanBeNull]
        private static T CastWithTypeCheckings<T>([CanBeNull] object from)
        {
            if(from is T res)
                return res;
            if(from == null && !typeof(T).IsValueType)
                return default;
            throw new ObjectPropertyExtractionException($"Can't assign item of type '{from?.GetType()}' to target of type '{typeof(T)}'");
        }

        [CanBeNull]
        private static T InitValue<T>([CanBeNull] T currentValue)
            where T : new()
        {
            if(typeof(T).IsValueType)
                return currentValue;
            if(currentValue != null)
                return currentValue;
            return new T();
        }

        [NotNull, ItemCanBeNull]
        private static T[] InitArray<T>([CanBeNull, ItemCanBeNull] T[] currentValue, int length)
            where T : new()
        {
            if(currentValue != null)
                return currentValue;
            return new T[length];
        }

        private static void InitDict<TKey, TValue>([NotNull] Dictionary<TKey, TValue> dict, [CanBeNull] object indexer, [NotNull] Func<TValue> createValue)
        {
            if(!(indexer is TKey realIndexer))
                throw new ObjectPropertyExtractionException($"Indexer type '{indexer?.GetType().ToString() ?? "NULL"}' does not match dictionary key type '{typeof(TKey)}'");
            if(indexer == null)
                throw new ObjectPropertyExtractionException("Can't use null as dict key");
            if(!dict.ContainsKey(realIndexer))
                dict[realIndexer] = createValue();
        }

        private static void InitPrimitiveDict<TKey, TValue>([NotNull] Dictionary<TKey, TValue> dict, [CanBeNull] object indexer)
        {
            InitDict(dict, indexer, () => default);
        }

        private static void InitClassDict<TKey, TValue>([NotNull] Dictionary<TKey, TValue> dict, [CanBeNull] object indexer)
            where TValue : new()
        {
            InitDict(dict, indexer, () => new TValue());
        }

        [NotNull]
        private static MethodInfo GetGenericMethod([NotNull] Type type, [NotNull] string name, [NotNull, ItemNotNull] params Type[] genericTypes)
        {
            return GetMethod(type, name).MakeGenericMethod(genericTypes);
        }

        [NotNull]
        private static MethodInfo GetMethod([NotNull] Type type, [NotNull] string name)
        {
            return type.GetMethod(name, BindingFlags.NonPublic | BindingFlags.Static) ?? throw new InvalidProgramStateException($"Method '{name}' not found in '{type}'");
        }
    }
}