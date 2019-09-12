using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using Excel.TemplateEngine.Exceptions;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.Helpers
{
    internal static class ObjectPropertiesExtractor
    {
        [CanBeNull]
        public static object ExtractChildObject([NotNull] object model, [NotNull] ExcelTemplatePath path)
        {
            var pathParts = path.PartsWithIndexers;
            var (succeed, result) = TryExtractChildObject(model, pathParts, 0);
            if (!succeed)
                throw new ObjectPropertyExtractionException($"Can't find path '{string.Join(".", pathParts)}' in model of type '{model.GetType()}'");

            return result;
        }

        private static (bool succeed, object result) TryExtractChildObject([NotNull] object model, [NotNull, ItemNotNull] string[] pathParts, int pathPartIndex)
        {
            if (pathPartIndex == pathParts.Length)
                return (true, model);

            if (!TryExtractDirectChild(model, pathParts[pathPartIndex], out var currentChild))
                return (false, null);

            if (currentChild == null)
                return (true, null);

            if (TemplateDescriptionHelper.IsArrayPathPart(pathParts[pathPartIndex]))
            {
                if (!TypeCheckingHelper.IsEnumerable(currentChild.GetType()))
                    throw new ObjectPropertyExtractionException($"Trying to extract enumerable from non-enumerable property {string.Join(".", pathParts)}");
                var resultList = new List<object>();
                foreach (var element in ((IEnumerable)currentChild).Cast<object>())
                {
                    if (element == null)
                        resultList.Add(null);
                    else
                    {
                        var (succeed, result) = TryExtractChildObject(element, pathParts, pathPartIndex + 1);
                        if (!succeed)
                            return (false, result);
                        resultList.Add(result);
                    }
                }
                return (true, resultList.ToArray());
            }

            return TryExtractChildObject(currentChild, pathParts, pathPartIndex + 1);
        }

        private static bool TryExtractDirectChild([NotNull] object model, [NotNull] string pathPart, out object child)
        {
            child = null;
            if (TemplateDescriptionHelper.IsCollectionAccessPathPart(pathPart))
            {
                var name = TemplateDescriptionHelper.GetCollectionAccessPathPartName(pathPart);
                var key = TemplateDescriptionHelper.GetCollectionAccessPathPartIndex(pathPart);
                if (!TryExtractCurrentChildPropertyInfo(model, name, out var collectionPropertyInfo))
                    return false;
                if (TypeCheckingHelper.IsDictionary(collectionPropertyInfo.PropertyType))
                {
                    var indexer = TemplateDescriptionHelper.ParseCollectionIndexerOrThrow(key, TypeCheckingHelper.GetDictionaryGenericTypeArguments(collectionPropertyInfo.PropertyType).keyType);
                    var dict = collectionPropertyInfo.GetValue(model, null);
                    if (dict == null)
                        return true;
                    child = ((IDictionary)dict)[indexer];
                    return true;
                }
                if (TypeCheckingHelper.IsIList(collectionPropertyInfo.PropertyType))
                {
                    var indexer = (int)TemplateDescriptionHelper.ParseCollectionIndexerOrThrow(key, typeof(int));
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

        private static bool TryExtractCurrentChildPropertyInfo([NotNull] object model, [NotNull] string pathPart, out PropertyInfo childPropertyInfo)
        {
            var propertyName = TemplateDescriptionHelper.GetPathPartName(pathPart);
            childPropertyInfo = model.GetType().GetProperty(propertyName);
            return childPropertyInfo != null;
        }

        [NotNull]
        private static PropertyInfo ExtractPropertyInfo([NotNull] Type type, [NotNull] string pathPart)
        {
            var propertyName = TemplateDescriptionHelper.GetPathPartName(pathPart);
            var childPropertyInfo = type.GetProperty(propertyName);
            if (childPropertyInfo == null)
                throw new ObjectPropertyExtractionException($"Property with name '{propertyName}' not found in type '{type}'");
            return childPropertyInfo;
        }

        [NotNull]
        public static Type ExtractChildObjectTypeFromPath([NotNull] Type modelType, [NotNull] ExcelTemplatePath path)
        {
            var currType = modelType;

            foreach (var part in path.PartsWithIndexers)
            {
                var childPropertyType = ExtractPropertyInfo(currType, part).PropertyType;

                if (TemplateDescriptionHelper.IsCollectionAccessPathPart(part))
                {
                    if (TypeCheckingHelper.IsDictionary(childPropertyType))
                    {
                        var (keyType, valueType) = TypeCheckingHelper.GetDictionaryGenericTypeArguments(childPropertyType);
                        TemplateDescriptionHelper.ParseCollectionIndexerOrThrow(TemplateDescriptionHelper.GetCollectionAccessPathPartIndex(part), keyType);
                        currType = valueType;
                    }
                    else if (TypeCheckingHelper.IsIList(childPropertyType))
                    {
                        TemplateDescriptionHelper.ParseCollectionIndexerOrThrow(TemplateDescriptionHelper.GetCollectionAccessPathPartIndex(part), typeof(int));
                        currType = TypeCheckingHelper.GetEnumerableItemType(childPropertyType);
                    }
                    else
                        throw new ObjectPropertyExtractionException($"Not supported collection type {childPropertyType}");
                }
                else if (TemplateDescriptionHelper.IsArrayPathPart(part))
                {
                    if (TypeCheckingHelper.IsIList(childPropertyType))
                        currType = TypeCheckingHelper.GetIListItemType(childPropertyType);
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
    }
}