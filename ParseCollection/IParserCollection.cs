using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection
{
    public interface IParserCollection
    {
        IClassParser GetClassParser();
        IEnumerableParser GetEnumerableParser(Type modelType);
        IAtomicValueParser GetAtomicValueParser();
        IFormValueParser GetFormValueParser(string formControlTypeName, Type valueType);
        IEnumerableMeasurer GetEnumerableMeasurer();
    }
}