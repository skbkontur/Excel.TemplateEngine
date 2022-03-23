using System;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection
{
    internal interface IParserCollection
    {
        IClassParser GetClassParser();
        IEnumerableParser GetEnumerableParser(Type modelType);
        LazyClassParser GetLazyClassParser();
        IAtomicValueParser GetAtomicValueParser();
        IFormValueParser GetFormValueParser(string formControlTypeName, Type valueType);
        IEnumerableMeasurer GetEnumerableMeasurer();
    }
}