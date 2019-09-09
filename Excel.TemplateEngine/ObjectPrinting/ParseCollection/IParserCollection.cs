using System;

using Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection
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