using System;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations;

using Vostok.Logging.Abstractions;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection
{
    internal class ParserCollection : IParserCollection
    {
        public ParserCollection(ILog logger)
        {
            this.logger = logger;
        }

        public IClassParser GetClassParser()
        {
            return new ClassParser(this, logger);
        }

        public IEnumerableParser GetEnumerableParser(Type modelType)
        {
            if (TypeCheckingHelper.IsEnumerable(modelType))
                return new EnumerableParser(this);
            throw new InvalidOperationException($"{modelType} is not IEnumerable");
        }

        public ListParser GetListParser(Type modelType)
        {
            if (TypeCheckingHelper.IsIList(modelType))
                return new ListParser();
            throw new InvalidOperationException($"{modelType} is not IEnumerable");
        }

        public IAtomicValueParser GetAtomicValueParser()
        {
            return new AtomicValueParser();
        }

        public IFormValueParser GetFormValueParser(string formControlTypeName, Type valueType)
        {
            if (formControlTypeName == "CheckBox" && valueType == typeof(bool))
                return new CheckBoxValueParser();
            if (formControlTypeName == "DropDown" && valueType == typeof(string))
                return new DropDownValueParser();
            throw new InvalidOperationException($"Unsupported pair of {nameof(formControlTypeName)} ({formControlTypeName}) and {nameof(valueType)} ({valueType}) for form controls");
        }

        public IEnumerableMeasurer GetEnumerableMeasurer()
        {
            return new EnumerableMeasurer(this);
        }

        private readonly ILog logger;
    }
}