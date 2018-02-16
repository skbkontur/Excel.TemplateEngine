using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public static class ObjectPropertiesExtractor
    {
        [CanBeNull]
        public static object ExtractChildObject([NotNull] object model, [NotNull] ExcelTemplateExpression expression)
        {
            return ExtractChildObject(model, expression.ChildObjectPath);
        }

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

            if(!TryExtractDirectChild(model, pathParts[pathPartIndex], out var currentChild))
                return (false, null);
            
            if (currentChild == null)
                return (true, null);
            
            if (TemplateDescriptionHelper.Instance.IsArrayPathPart(pathParts[pathPartIndex]))
            {
                if (!TypeCheckingHelper.Instance.IsEnumerable(currentChild.GetType()))
                    throw new ObjectPropertyExtractionException($"Trying to extract enumerable from non-enumerable property {string.Join(".", pathParts)}");
                var resultList = new List<object>();
                foreach (var element in ((IEnumerable)currentChild).Cast<object>())
                {
                    if (element == null)
                        resultList.Add(null);
                    else
                    {
                        var (succeed, result) = TryExtractChildObject(element, pathParts, pathPartIndex + 1);
                        if(!succeed)
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
            if (TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(pathPart))
            {
                var name = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartName(pathPart);
                var key = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(pathPart);
                if (!TryExtractCurrentChildPropertyInfo(model, name, out var collectionPropertyInfo))
                    return false;
                if (TypeCheckingHelper.Instance.IsDictionary(collectionPropertyInfo.PropertyType))
                {
                    var indexer = TemplateDescriptionHelper.ParseCollectionIndexer(key, TypeCheckingHelper.Instance.GetDictionaryKeyType(collectionPropertyInfo.PropertyType));
                    var dict = collectionPropertyInfo.GetValue(model, null);
                    if (dict == null)
                        return true;
                    child = ((IDictionary)dict)[indexer];
                    return true;
                }
                if (TypeCheckingHelper.Instance.IsIList(collectionPropertyInfo.PropertyType))
                {
                    var indexer = (int)TemplateDescriptionHelper.ParseCollectionIndexer(key, typeof(int));
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
            var propertyName = TemplateDescriptionHelper.Instance.GetPathPartName(pathPart);
            childPropertyInfo = model.GetType().GetProperty(propertyName);
            return childPropertyInfo != null;
        }

        [NotNull]
        private static PropertyInfo ExtractPropertyInfo([NotNull] Type type, [NotNull] string pathPart)
        {
            var propertyName = TemplateDescriptionHelper.Instance.GetPathPartName(pathPart);
            var childPropertyInfo = type.GetProperty(propertyName);
            if (childPropertyInfo == null)
                throw new ObjectPropertyExtractionException($"Property with name '{propertyName}' not found in type '{type}'");
            return childPropertyInfo;
        }

        [NotNull]
        public static Type ExtractChildObjectTypeFromPath([NotNull] Type modelType, [NotNull] ExcelTemplatePath path)
        {
            var currType = modelType;

            foreach(var part in path.PartsWithIndexers)
            {
                var childPropertyType = ExtractPropertyInfo(currType, part).PropertyType;

                if(TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(part))
                {
                    if(TypeCheckingHelper.Instance.IsDictionary(childPropertyType))
                    {
                        TemplateDescriptionHelper.ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), TypeCheckingHelper.Instance.GetDictionaryKeyType(childPropertyType));
                        currType = TypeCheckingHelper.Instance.GetDictionaryValueType(childPropertyType);
                    }
                    else if(TypeCheckingHelper.Instance.IsIList(childPropertyType))
                    {
                        TemplateDescriptionHelper.ParseCollectionIndexer(TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(part), typeof(int));
                        currType = TypeCheckingHelper.Instance.GetEnumerableItemType(childPropertyType);
                    }
                    else
                        throw new ObjectPropertyExtractionException($"Not supported collection type {childPropertyType}");
                }
                else if(TemplateDescriptionHelper.Instance.IsArrayPathPart(part))
                {
                    if(TypeCheckingHelper.Instance.IsIList(childPropertyType))
                        currType = TypeCheckingHelper.Instance.GetIListItemType(childPropertyType);
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