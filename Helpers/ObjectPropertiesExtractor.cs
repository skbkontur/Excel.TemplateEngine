using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;

using JetBrains.Annotations;

using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.Helpers
{
    public sealed class ObjectPropertiesExtractor
    {
        private ObjectPropertiesExtractor()
        {
        }

        [CanBeNull]
        public object ExtractChildObject(object model, string expression)
        {
            if(!(TemplateDescriptionHelper.Instance.IsCorrectValueDescription(expression) || TemplateDescriptionHelper.Instance.IsCorrectFormValueDescription(expression)) ||
               model == null)
                return null;

            var descriptionParts = TemplateDescriptionHelper.Instance.GetDescriptionParts(expression);
            var pathParts = descriptionParts[2].Split('.');

            return ExtractChildObject(model, pathParts);
        }

        public static Type ExtractChildObjectType([NotNull] object model, [NotNull] string expression)
        {
            return ExtractChildObjectTypeFromPath(model, ExtractCleanChildObjectPath(expression));
        }

        public static Type ExtractChildObjectTypeFromPath([NotNull] object model, [NotNull] string path)
        {
            var currType = model.GetType();
            foreach (var part in path.Split('.'))
            {
                if (TypeCheckingHelper.Instance.IsDictionary(currType))
                    currType = TypeCheckingHelper.Instance.GetDictionaryValueType(currType);
                else if (TypeCheckingHelper.Instance.IsEnumerable(currType))
                    currType = TypeCheckingHelper.Instance.GetEnumerableItemType(currType);
                
                if(TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(part)) // todo (mpivko, 17.12.2017): assert IsCollectionAccessPathPartName == true only for enumerables and dicts
                {
                    var cleanPart = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartName(part);
                    var property = currType.GetProperty(cleanPart);
                    if (property == null)
                        throw new Exception($"Type '{currType}' has no property '{cleanPart}'");
                    if (TypeCheckingHelper.Instance.IsDictionary(property.PropertyType))
                        currType = TypeCheckingHelper.Instance.GetDictionaryValueType(property.PropertyType);
                    else if (TypeCheckingHelper.Instance.IsEnumerable(property.PropertyType))
                        currType = TypeCheckingHelper.Instance.GetEnumerableItemType(property.PropertyType);
                    else
                        throw new Exception(); // todo (mpivko, 17.12.2017): 
                }
                else
                {
                    var cleanPart = part.Replace("[]", "");
                    var property = currType.GetProperty(cleanPart);
                    if (property == null)
                        throw new Exception($"Type '{currType}' has no property '{cleanPart}'");
                    currType = property.PropertyType; //TODO mpivko strange
                }
            }
            return currType;//TODO mpivko maybe there is simplier way
            //return Instance.ExtractChildObject(model, expression).GetType(); 
        }

        public static (string, string) SplitForEnumerableExpansion([NotNull] string expression)
        {
            var parts = ExtractChildObjectPath(expression).Split('.');
            if (!parts.Any(x => x.Contains("[]")))
                throw new Exception(); // todo (mpivko, 08.12.2017): 
            var firstPartLen = parts.TakeWhile(x => !x.EndsWith("[]")).Count() + 1;
            return (string.Join(".", parts.Take(firstPartLen)), string.Join(".", parts.Skip(firstPartLen)));
        }

        public static bool NeedEnumerableExpansion([NotNull] string expression)
        {
            return ExtractChildObjectPath(expression).Contains("[]");
        }

        public static string ExtractChildObjectPath([NotNull] string expression)
        {
            return expression.Split(':')[2];
        }

        public static string ExtractCleanChildObjectPath([NotNull] string expression)
        {
            return ExtractChildObjectPath(expression).Replace("[]", "");
        }

        [NotNull]
        public static Action<object> ExtractChildObjectSetter([NotNull] object model, [NotNull] string expression)
        {
            if (!TemplateDescriptionHelper.Instance.IsCorrectAbstractValueDescription(expression))
                throw new InvalidProgramStateException($"Invalid description {expression}");

            var descriptionParts = TemplateDescriptionHelper.Instance.GetDescriptionParts(expression);
            var pathParts = descriptionParts[2].Split('.');
            
            return ExtractChildModelSetter(model, pathParts, 0);
        }
        
        public static ObjectPropertiesExtractor Instance { get; } = new ObjectPropertiesExtractor();

        private static bool TryExtractCurrentChildPropertyInfo(object model, string pathPart, out PropertyInfo childPropertyInfo)
        {
            var propertyName = TemplateDescriptionHelper.Instance.GetPathPartName(pathPart);
            childPropertyInfo = model.GetType()
                                    .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                                    .FirstOrDefault(property => property.Name == propertyName);
            return childPropertyInfo != null;
        }

        [NotNull]
        private static PropertyInfo ExtractCurrentChildPropertyInfo(object model, string pathPart)
        {
            var propertyName = TemplateDescriptionHelper.Instance.GetPathPartName(pathPart);
            var childPropertyInfo = model.GetType()
                                     .GetProperties(BindingFlags.Instance | BindingFlags.Public)
                                     .FirstOrDefault(property => property.Name == propertyName);
            if (childPropertyInfo == null)
                throw new InvalidProgramStateException($"Property with name '{propertyName}' not found in '{model.GetType()}'");
            return childPropertyInfo;
        }

        private static bool TryExtractCurrentChild(object model, string pathPart, out object child)
        {
            child = null;
            if (TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(pathPart))
            {
                var name = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartName(pathPart);
                var key = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(pathPart);
                if (!TryExtractCurrentChildPropertyInfo(model, name, out var dictPropertyInfo))
                    return false;
                if(!TypeCheckingHelper.Instance.IsDictionary(dictPropertyInfo.PropertyType))
                    throw new Exception($"Unexpected child type: expected dictionary (pathPath='{pathPart}'), but model is {model.GetType()}");
                if(key.StartsWith("\"") && key.EndsWith("\"")) // todo (mpivko, 19.12.2017): copypaste from printer
                {
                    child = ((IDictionary)dictPropertyInfo.GetValue(model, null))[key.Substring(1, key.Length - 2)];
                }
                else if(int.TryParse(key, out var keyInt))
                {
                    child = ((IDictionary)dictPropertyInfo.GetValue(model, null))[keyInt];
                }
                else
                {
                    throw new NotSupportedException($"Can't parse '{key}' as a dictionary key");
                }
                return true;
            }
            if (!TryExtractCurrentChildPropertyInfo(model, pathPart, out var propertyInfo))
                return false;
            child = propertyInfo.GetValue(model, null);
            return true;
        }

        private static void ExtractChildObject(object model, string[] pathParts, int pathPartIndex, List<object> result)
        {
            if(pathPartIndex == pathParts.Length)
            {
                result.Add(model);
                return;
            }

            if (!TryExtractCurrentChild(model, pathParts[pathPartIndex], out var currentChild))
                return;
            
            if (currentChild == null)
            {
                result.Add(null);
                return;
            }

            if(TypeCheckingHelper.Instance.IsEnumerable(currentChild.GetType()))
            {
                foreach(var element in ((IEnumerable)currentChild).Cast<object>())
                    ExtractChildObject(element, pathParts, pathPartIndex + 1, result);
                return;
            }

            ExtractChildObject(currentChild, pathParts, pathPartIndex + 1, result);
        }

        private static object ExtractChildObject(object model, string[] pathParts)
        {
            var result = new List<object>();
            ExtractChildObject(model, pathParts, 0, result);

            if(result.Count == 0)
                return null;

            if(pathParts.Any(TemplateDescriptionHelper.Instance.IsArrayPathPart))
                return result.ToArray();
            return result.Single();
        }

        private static Action<object> ExtractChildModelSetter(object model, string[] pathParts, int pathPartIndex)
        {
            var childPropertyInfo = ExtractCurrentChildPropertyInfo(model, pathParts[pathPartIndex]);

            if (TemplateDescriptionHelper.Instance.IsCollectionAccessPathPart(pathParts[pathPartIndex]))
            {
                var collectionAccessPart = TemplateDescriptionHelper.Instance.GetCollectionAccessPathPartIndex(pathParts[pathPartIndex]);
                if(collectionAccessPart.StartsWith("\""))
                {
                    if (TypeCheckingHelper.Instance.IsDictionary(childPropertyInfo.PropertyType) && typeof(string).IsAssignableFrom(TypeCheckingHelper.Instance.GetDictionaryKeyType(childPropertyInfo.PropertyType)))
                    {
                        var stringIndex = collectionAccessPart.Substring(1, collectionAccessPart.Length - 2);
                        if (stringIndex.Contains("\""))
                            throw new NotSupportedException("Escaping on keys is not supported");
                        return value =>
                            {
                                if(childPropertyInfo.GetValue(model) == null)
                                {
                                    childPropertyInfo.SetValue(model, Activator.CreateInstance(childPropertyInfo.PropertyType));
                                }
                                if (!(childPropertyInfo.GetValue(model) is IDictionary dictionaryValue))
                                    throw new InvalidProgramStateException($"Expected IDictionary, but got {value.GetType()} when trying to set {model.GetType()}.{childPropertyInfo.Name}");

                                if(pathPartIndex + 1 == pathParts.Length)
                                {
                                    dictionaryValue[stringIndex] = value;
                                }
                                else
                                {
                                    if (!dictionaryValue.Contains(stringIndex))
                                        dictionaryValue[stringIndex] = Activator.CreateInstance(TypeCheckingHelper.Instance.GetDictionaryValueType(childPropertyInfo.PropertyType));
                                    var childSetter = ExtractChildModelSetter(dictionaryValue[stringIndex], pathParts, pathPartIndex + 1);
                                    childSetter(dictionaryValue[stringIndex]);
                                }
                            };
                    }
                    throw new Exception(); // todo (mpivko, 17.12.2017): 
                }
                if (int.TryParse(collectionAccessPart, out var index))
                {
                    if (TypeCheckingHelper.Instance.IsDictionary(childPropertyInfo.PropertyType) && typeof(int).IsAssignableFrom(TypeCheckingHelper.Instance.GetDictionaryKeyType(childPropertyInfo.PropertyType)))
                    {
                        return value =>
                            {
                                // todo (mpivko, 17.12.2017): this code is the same as for strings
                                if (childPropertyInfo.GetValue(model) == null)
                                {
                                    childPropertyInfo.SetValue(model, Activator.CreateInstance(childPropertyInfo.PropertyType));
                                }
                                if (!(childPropertyInfo.GetValue(model) is IDictionary dictionaryValue))
                                    throw new InvalidProgramStateException($"Expected IDictionary, but got {value.GetType()} when trying to set {model.GetType()}.{childPropertyInfo.Name}");

                                if (pathPartIndex + 1 == pathParts.Length)
                                {
                                    dictionaryValue[index] = value;
                                }
                                else
                                {
                                    if (!dictionaryValue.Contains(index))
                                        dictionaryValue[index] = Activator.CreateInstance(TypeCheckingHelper.Instance.GetDictionaryValueType(childPropertyInfo.PropertyType));
                                    var childSetter = ExtractChildModelSetter(dictionaryValue[index], pathParts, pathPartIndex + 1);
                                    childSetter(dictionaryValue[index]);
                                }
                            };
                    }
                    if(TypeCheckingHelper.Instance.IsEnumerable(childPropertyInfo.PropertyType))
                    {
                        return value =>
                            {
                                // todo (mpivko, 17.12.2017): this code is very simular to one used for strings
                                if (childPropertyInfo.GetValue(model) == null)
                                {
                                    // todo (mpivko, 17.12.2017): there is a bug here: we can't create array here, for example, because we don't know it's size
                                    throw new NotImplementedException();
                                    childPropertyInfo.SetValue(model, Activator.CreateInstance(childPropertyInfo.PropertyType));
                                }
                                if (!(childPropertyInfo.GetValue(model) is IList listValue))
                                    throw new InvalidProgramStateException($"Expected IList, but got {value.GetType()} when trying to set {model.GetType()}.{childPropertyInfo.Name}");

                                if (pathPartIndex + 1 == pathParts.Length)
                                {
                                    listValue[index] = value;
                                }
                                else
                                {
                                    if (listValue.Count <= index)
                                        throw new OverflowException(); // todo (mpivko, 17.12.2017): better exception
                                    if (listValue[index] == null)
                                        listValue[index] = Activator.CreateInstance(TypeCheckingHelper.Instance.GetDictionaryValueType(childPropertyInfo.PropertyType));
                                    var childSetter = ExtractChildModelSetter(listValue[index], pathParts, pathPartIndex + 1);
                                    childSetter(listValue[index]);
                                }
                            };
                    }
                    throw new Exception(); // todo (mpivko, 17.12.2017): 
                }
                throw new NotSupportedException($"Unknown index type. Index: '{collectionAccessPart}'");
            }

            if (pathPartIndex + 1 == pathParts.Length)
            {
                return value => childPropertyInfo.SetValue(model, value);
            }

            if (TypeCheckingHelper.Instance.IsEnumerable(childPropertyInfo.PropertyType))
            {
                Action<object> childIEnumerableSetter = value =>
                    {
                        if(!(value is IEnumerable enumerableValue))
                            throw new InvalidProgramStateException($"Expected IEnumerable, but got {value.GetType()} when trying to set {model.GetType()}.{childPropertyInfo.Name}");
                        var listValue = enumerableValue.Cast<object>().ToList();
                        if (!typeof(IList).IsAssignableFrom(childPropertyInfo.PropertyType))
                            throw new NotSupportedException("Only IList ienumerables are supported");
                        var subEnumerable = (IList)childPropertyInfo.GetValue(model);
                        if (subEnumerable == null)
                        {
                            subEnumerable = (IList)Activator.CreateInstance(childPropertyInfo.PropertyType, listValue.Count);
                            childPropertyInfo.SetValue(model, subEnumerable);
                        }

                        for(var i = 0; i < listValue.Count; i++)
                        {
                            if (pathPartIndex + 1 == pathParts.Length)
                            {
                                subEnumerable[i] = listValue[i];
                            }
                            else
                            {
                                if(subEnumerable[i] == null)
                                    subEnumerable[i] = Activator.CreateInstance(TypeCheckingHelper.Instance.GetEnumerableItemType(childPropertyInfo.PropertyType));
                                var childSetter = ExtractChildModelSetter(subEnumerable[i], pathParts, pathPartIndex + 1);
                                childSetter(listValue[i]);
                            }
                        }
                    };
                return childIEnumerableSetter;
            }

            /*if(childPropertyInfo.GetValue(model) == null)
            {
                var childConstructor = childPropertyInfo.PropertyType.GetConstructor(new Type[0]);
                if (childConstructor == null)
                    throw new InvalidProgramStateException($"Failed to parse xlsx: {childPropertyInfo.PropertyType} has no parameterless constructor");
                childPropertyInfo.SetValue(model, childConstructor.Invoke(new object[0]));
            }*/
            var subValue = childPropertyInfo.GetValue(model);
            if (subValue == null)
            {
                subValue = Activator.CreateInstance(childPropertyInfo.PropertyType);
                childPropertyInfo.SetValue(model, subValue);
            }
            var childPropertySetter = ExtractChildModelSetter(subValue, pathParts, pathPartIndex + 1);

            return value => childPropertySetter(value);
        }
    }
}