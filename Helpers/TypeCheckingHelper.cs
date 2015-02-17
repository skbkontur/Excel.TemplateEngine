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

        public static TypeCheckingHelper Instance { get { return instance; } }

        private static bool IsEnumerableDirectly(Type type)
        {
            return type == typeof(IEnumerable) || (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(IEnumerable<>));
        }

        private static readonly TypeCheckingHelper instance = new TypeCheckingHelper();
    }
}