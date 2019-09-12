using System;

using Excel.TemplateEngine.Exceptions;
using Excel.TemplateEngine.ObjectPrinting.Helpers;
using Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers;
using Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations;

using Vostok.Logging.Abstractions;

namespace Excel.TemplateEngine.ObjectPrinting.ParseCollection
{
    public class ParserCollection : IParserCollection
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
            throw new ExcelTemplateEngineException($"{modelType} is not IEnumerable");
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
            throw new ExcelTemplateEngineException($"Unsupported pair of {nameof(formControlTypeName)} ({formControlTypeName}) and {nameof(valueType)} ({valueType}) for form controls");
        }

        public IEnumerableMeasurer GetEnumerableMeasurer()
        {
            return new EnumerableMeasurer(this);
        }

        private readonly ILog logger;
    }
}