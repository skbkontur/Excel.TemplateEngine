using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;

using SkbKontur.Excel.TemplateEngine.Exceptions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers
{
    internal static class TypeCheckingHelper
    {
        public static bool IsEnumerable(Type type)
            => type != typeof(string) &&
               (IsEnumerableDirectly(type) || type.GetInterfaces().Any(IsEnumerableDirectly));

        public static bool IsDictionary(Type type) 
            => IsDictionaryDirectly(type) || type.GetInterfaces().Any(IsDictionaryDirectly);

        public static bool IsIList(Type type) 
            => type != typeof(string) && (IsIListDirectly(type) || type.GetInterfaces().Any(IsIListDirectly));

        public static bool IsNullable(Type type) 
            => Nullable.GetUnderlyingType(type) != null;

        public static Type GetEnumerableItemType(Type type) 
            => GetImplementedEnumerableInterface(type).GetGenericArguments().SingleOrDefault() ?? typeof(object);

        public static Type GetIListItemType(Type type) 
            => GetImplementedIListInterface(type).GetGenericArguments().SingleOrDefault() ?? typeof(object);

        public static (Type keyType, Type valueType) GetDictionaryGenericTypeArguments(Type type)
        {
            if (!IsDictionary(type))
                throw new ExcelTemplateEngineException($"{nameof(type)} ({type}) should implement IDictionary<,> or IDictionary");
            var genericArguments = GetImplementedDictionaryInterface(type).GetGenericArguments();
            if (!genericArguments.Any())
                return (typeof(object), typeof(object));
            if (genericArguments.Length != 2)
                throw new ExcelTemplateEngineException($"Dict can have only 0 or 2 generic arguments, but here is {genericArguments.Length} of them ({string.Join(", ", genericArguments.Select(x => x.ToString()))}). Type is '{type}'.");
            return (genericArguments[0], genericArguments[1]);
        }

        private static Type GetImplementedEnumerableInterface(Type type)
        {
            if (type == typeof(string))
                return null;
            if (IsGenericEnumerableDirectly(type))
                return type;
            return type.GetInterfaces().FirstOrDefault(IsGenericEnumerableDirectly) ?? type.GetInterfaces().FirstOrDefault(IsEnumerableDirectly);
        }

        private static Type GetImplementedDictionaryInterface(Type type)
        {
            if (IsGenericDictionaryDirectly(type))
                return type;
            return type.GetInterfaces().FirstOrDefault(IsGenericDictionaryDirectly) ?? type.GetInterfaces().FirstOrDefault(IsDictionaryDirectly);
        }

        private static Type GetImplementedIListInterface(Type type)
        {
            if (IsGenericIListDirectly(type))
                return type;
            return type.GetInterfaces().FirstOrDefault(IsGenericIListDirectly) ?? type.GetInterfaces().FirstOrDefault(IsIListDirectly);
        }

        private static bool IsGenericEnumerableDirectly(Type type) 
            => type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IEnumerable<>);

        private static bool IsGenericDictionaryDirectly(Type type) 
            => type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IDictionary<,>);

        private static bool IsEnumerableDirectly(Type type) 
            => type == typeof(IEnumerable) || IsGenericEnumerableDirectly(type);

        private static bool IsDictionaryDirectly(Type type) 
            => type == typeof(IDictionary) || IsGenericDictionaryDirectly(type);

        private static bool IsIListDirectly(Type type) 
            => type == typeof(IList) || IsGenericIListDirectly(type);

        private static bool IsGenericIListDirectly(Type type) 
            => type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IList<>);
    }
}