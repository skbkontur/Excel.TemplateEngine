using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public sealed class TypeCheckingHelper
    {
        private TypeCheckingHelper()
        {
        }

        public bool IsEnumerable(Type type)
        {
            return type != typeof(string) &&
                   (IsEnumerableDirectly(type) || type.GetInterfaces().Any(IsEnumerableDirectly));
        }

        public bool IsDictionary(Type type)
        {
            return IsDictionaryDirectly(type) || type.GetInterfaces().Any(IsDictionaryDirectly);
        }

        public Type GetEnumerableItemType(Type type)
        {
            return GetImplementedEnumerableInterface(type).GetGenericArguments().SingleOrDefault() ?? typeof(object); // todo (mpivko, 17.12.2017): suppport non-generic enumerables
        }

        public Type GetDictionaryKeyType(Type type)
        {
            return GetDictionaryGenericTypeArguments(type).keyType;
        }

        public Type GetDictionaryValueType(Type type)
        {
            return GetDictionaryGenericTypeArguments(type).valueType;
        }

        private (Type keyType, Type valueType) GetDictionaryGenericTypeArguments(Type type)
        {
            if (!IsDictionary(type))
                throw new ArgumentException($"{nameof(type)} ({type}) should implement IDictionary<,> or IDictionary");
            var genericArguments = GetImplementedDictionaryInterface(type).GetGenericArguments();
            if (!genericArguments.Any())
                return (typeof(object), typeof(object)); // todo (mpivko, 17.12.2017): suport non-generic dictionaries 
            return (genericArguments.First(), genericArguments.Skip(1).First());
        }

        private Type GetImplementedEnumerableInterface(Type type)
        {
            if (type == typeof(string))
                return null;
            if (IsGenericEnumerableDirectly(type))
                return type;
            return type.GetInterfaces().FirstOrDefault(IsGenericEnumerableDirectly) ?? type.GetInterfaces().FirstOrDefault(IsEnumerableDirectly);
        }

        private Type GetImplementedDictionaryInterface(Type type)
        {
            if (IsGenericDictionaryDirectly(type))
                return type;
            return type.GetInterfaces().FirstOrDefault(IsGenericDictionaryDirectly) ?? type.GetInterfaces().FirstOrDefault(IsDictionaryDirectly);
        }

        public static TypeCheckingHelper Instance { get { return instance; } }

        private static bool IsGenericEnumerableDirectly(Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IEnumerable<>);
        }

        private static bool IsGenericDictionaryDirectly(Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IDictionary<,>);
        }

        private static bool IsEnumerableDirectly(Type type)
        {
            return type == typeof(IEnumerable) || IsGenericEnumerableDirectly(type);
        }

        private static bool IsDictionaryDirectly(Type type)
        {
            return type == typeof(IDictionary) || IsGenericDictionaryDirectly(type);
        }

        private static readonly TypeCheckingHelper instance = new TypeCheckingHelper();
    }
}