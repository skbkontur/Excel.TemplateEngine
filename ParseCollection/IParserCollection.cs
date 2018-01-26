using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection
{
    public interface IParserCollection
    {
        IClassParser GetClassParser(Type modelType);
        IEnumerableParser GetEnumerableParser(Type modelType);
        IAtomicValueParser GetAtomicValueParser(Type valueType);
        IFormValueParser GetFormValueParser(string formControlTypeName, Type valueType);
    }
}