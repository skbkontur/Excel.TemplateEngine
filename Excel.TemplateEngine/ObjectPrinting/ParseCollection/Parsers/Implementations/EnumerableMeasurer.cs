using System;
using System.Collections.Generic;

using SkbKontur.Excel.TemplateEngine.Exceptions;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations
{
    internal class EnumerableMeasurer : IEnumerableMeasurer
    {
        public EnumerableMeasurer(ParserCollection parserCollection)
        {
            this.parserCollection = parserCollection;
        }

        public int GetLength(ITableParser tableParser, Type modelType, IEnumerable<ICell> primaryParts)
        {
            var parserState = new List<(ICell cell, Type itemType)>();

            foreach (var primaryPart in primaryParts)
            {
                var childModelPath = ExcelTemplatePath.FromRawExpression(primaryPart.StringValue);
                var childModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, childModelPath);

                parserState.Add((primaryPart, childModelType));
            }

            for (var i = 0; i <= ParsingParameters.MaxEnumerableLength; i++)
            {
                var parsed = false;
                foreach (var (cell, type) in parserState)
                {
                    tableParser.PushState(cell.CellPosition.Add(new ObjectSize(0, i)));
                    if (TextValueParser.TryParse(tableParser.GetCurrentCellText(), type, out var result) && result != null)
                        parsed = true;
                    tableParser.PopState();
                }
                if (!parsed)
                    return i;
            }
            throw new EnumerableTooLongException(ParsingParameters.MaxEnumerableLength);
        }

        private readonly ParserCollection parserCollection;
    }
}