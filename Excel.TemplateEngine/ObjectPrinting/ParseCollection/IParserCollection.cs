using System;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection
{
    internal interface IParserCollection
    {
        [NotNull]
        IClassParser GetClassParser();

        [NotNull]
        IEnumerableParser GetEnumerableParser(Type modelType);

        [NotNull]
        LazyClassParser GetLazyClassParser();

        [NotNull]
        IFormValueParser GetFormValueParser(string formControlTypeName, Type valueType);

        [NotNull]
        IEnumerableMeasurer GetEnumerableMeasurer();
    }
}