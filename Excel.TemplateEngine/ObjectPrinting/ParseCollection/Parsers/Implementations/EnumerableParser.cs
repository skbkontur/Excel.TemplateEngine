using System;
using System.Collections.Generic;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations
{
    internal class EnumerableParser : IEnumerableParser
    {
        public EnumerableParser(IParserCollection parserCollection)
        {
            this.parserCollection = parserCollection;
        }

        [NotNull]
        public List<object> Parse([NotNull] ITableParser tableParser, [NotNull] Type modelType, int count, [NotNull] Action<string, string> addFieldMapping)
        {
            if (count < 0)
                throw new InvalidOperationException($"Count should be positive ({count} found)");
            if (count > ParsingParameters.MaxEnumerableLength)
                throw new InvalidOperationException($"Lists longer than {ParsingParameters.MaxEnumerableLength} are not supported");

            var result = new List<object>();
            for (var i = 0; i < count; i++)
            {
                if (i != 0)
                    tableParser.MoveToNextLayer();

                tableParser.PushState();

                if (!TextValueParser.TryParse(tableParser.GetCurrentCellText(), modelType, out var item) || item == null)
                    item = GetDefault(modelType);

                addFieldMapping($"[{i}]", tableParser.CurrentState.Cursor.CellReference);
                result.Add(item);
                tableParser.PopState();
            }

            return result;
        }

        private static object GetDefault(Type type)
        {
            if (type.IsValueType)
                return Activator.CreateInstance(type);
            return null;
        }

        private readonly IParserCollection parserCollection;
    }
}