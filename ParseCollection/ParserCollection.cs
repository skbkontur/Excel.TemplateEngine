using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection
{
    public class ParserCollection : IParserCollection
    {
        public IClassParser GetClassParser(Type modelType)
        {
            return new ClassParser(this);
        }

        public IEnumerableParser GetEnumerableParser(Type modelType)
        {
            if (TypeCheckingHelper.Instance.IsEnumerable(modelType))
                return new EnumerableParser(this);
            throw new ArgumentException();
        }

        public IAtomicValueParser GetAtomicValueParser(Type valueType)
        {
            return new AtomicValueParser();
        }

        public IFormValueParser GetFormValueParser(string formControlTypeName, Type valueType)
        {
            if (formControlTypeName == "CheckBox" && valueType == typeof(bool))
                return new CheckBoxValueParser();
            if (formControlTypeName == "DropDown" && valueType == typeof(string))
                return new DropDownValueParser();
            throw new NotSupportedExcelSerializationException($"Unsupported pair of {nameof(formControlTypeName)} ({formControlTypeName}) and {nameof(valueType)} ({valueType}) for form controls");
        }
    }
}