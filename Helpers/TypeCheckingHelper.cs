using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using SKBKontur.Catalogue.Objects;

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

        public bool IsIList(Type type)
        {
            return type != typeof(string) && (IsIListDirectly(type) || type.GetInterfaces().Any(IsIListDirectly));
        }

        public bool IsNullable(Type type)
        {
            return Nullable.GetUnderlyingType(type) != null;
        }

        public Type GetEnumerableItemType(Type type)
        {
            return GetImplementedEnumerableInterface(type).GetGenericArguments().SingleOrDefault() ?? typeof(object);
        }

        public Type GetDictionaryKeyType(Type type)
        {
            return GetDictionaryGenericTypeArguments(type).keyType;
        }

        public Type GetDictionaryValueType(Type type)
        {
            return GetDictionaryGenericTypeArguments(type).valueType;
        }

        public Type GetIListItemType(Type type)
        {
            return GetImplementedIListInterface(type).GetGenericArguments().SingleOrDefault() ?? typeof(object);
        }

        private (Type keyType, Type valueType) GetDictionaryGenericTypeArguments(Type type)
        {
            if (!IsDictionary(type))
                throw new ArgumentException($"{nameof(type)} ({type}) should implement IDictionary<,> or IDictionary");
            var genericArguments = GetImplementedDictionaryInterface(type).GetGenericArguments();
            if (!genericArguments.Any())
                return (typeof(object), typeof(object));
            if (genericArguments.Length != 2)
                throw new InvalidProgramStateException($"Dict can have only 0 or 2 generic arguments, but here is {genericArguments.Length} of them ({string.Join(", ", genericArguments.Select(x => x.ToString()))}). Type is '{type}'.");
            return (genericArguments[0], genericArguments[1]);
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

        private Type GetImplementedIListInterface(Type type)
        {
            if (IsGenericIListDirectly(type))
                return type;
            return type.GetInterfaces().FirstOrDefault(IsGenericIListDirectly) ?? type.GetInterfaces().FirstOrDefault(IsIListDirectly);
        }

        public static TypeCheckingHelper Instance { get; } = new TypeCheckingHelper();

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

        private static bool IsIListDirectly(Type type)
        {
            return type == typeof(IList) || IsGenericIListDirectly(type);
        }

        private static bool IsGenericIListDirectly(Type type)
        {
            return type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IList<>);
        }
    }
}