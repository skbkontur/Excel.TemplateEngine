using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection
{
    public class ParserCollection : IParserCollection
    {
        public ParserCollection(ITemplateCollection templateCollection)
        {
            this.templateCollection = templateCollection;
        }

        public IClassParser GetParser(Type modelType)
        {
            //TODO mpivko avoid calling it with atomic values
            return new ClassParser(templateCollection, this);
        }

        public IEnumerableParser GetEnumerableParser(Type modelType)
        {
            //TODO mpivko avoid calling it with not enumerable
            if (TypeCheckingHelper.Instance.IsEnumerable(modelType))
                return new EnumerableParser(this);
            throw new ArgumentException();
        }

        public IAtomicValueParser GetAtomicValueParser(Type valueType)
        {
            if (valueType == typeof(string))
                return new StringParser();
            if (valueType == typeof(int))
                return new IntParser();
            if (valueType == typeof(decimal))
                return new DecimalParser();
            if (valueType == typeof(double))
                return new DoubleParser();
            if (valueType == typeof(long))
                return new LongParser();
            if (valueType == typeof(int?))
                return new IntParser(nullable : true);
            if (valueType == typeof(decimal?))
                return new DecimalParser(nullable : true);
            if (valueType == typeof(double?))
                return new DoubleParser(nullable : true);
            if (valueType == typeof(double?))
                return new LongParser(nullable: true);
            throw new NotSupportedException($"{valueType} is not a supported atomic value");
        }

        public IFormValueParser GetFormValueParser(Type valueType)
        {
            // todo (mpivko, 15.12.2017): maybe it's better not to derive Parser from valueType but use first part of value description instead ("this_one::")
            if (valueType == typeof(bool))
                return new CheckBoxValueParser();
            if (valueType == typeof(string))
                return new DropDownValueParser();
            throw new NotSupportedException($"{valueType} is not a supported atomic value");
        }

        private readonly ITemplateCollection templateCollection;
    }
}