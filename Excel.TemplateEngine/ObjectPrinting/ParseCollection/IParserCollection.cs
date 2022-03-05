using System;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection
{
    internal interface IParserCollection
    {
        IClassParser GetClassParser();
        IEnumerableParser GetEnumerableParser(Type modelType);
        ListParser GetListParser(Type modelType);
        IAtomicValueParser GetAtomicValueParser();
        IFormValueParser GetFormValueParser(string formControlTypeName, Type valueType);
        IEnumerableMeasurer GetEnumerableMeasurer();
    }
}