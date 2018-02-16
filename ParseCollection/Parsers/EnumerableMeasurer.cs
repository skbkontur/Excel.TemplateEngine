using System;
using System.Collections.Generic;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableParser;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ParseCollection.Parsers
{
    public class EnumerableMeasurer : IEnumerableMeasurer
    {
        public EnumerableMeasurer(ParserCollection parserCollection)
        {
            this.parserCollection = parserCollection;
        }

        public int GetLength(ITableParser tableParser, Type modelType, IEnumerable<ICell> primaryParts)
        {
            var parserState = new List<(IAtomicValueParser parser, ICell cell, Type itemType)>();

            foreach(var primaryPart in primaryParts)
            {
                tableParser.PushState(primaryPart.CellPosition, new Styler(primaryPart));

                var childModelPath = new ExcelTemplateExpression(primaryPart.StringValue).ChildObjectPath;
                var childModelType = ObjectPropertiesExtractor.ExtractChildObjectTypeFromPath(modelType, childModelPath);
                
                var parser = parserCollection.GetAtomicValueParser(childModelType);

                parserState.Add((parser, primaryPart, childModelType));
            }
            
            for (var i = 0; i <= maxEnumerableLength; i++)
            {
                var parsed = false;
                foreach(var (parser, cell, type) in parserState)
                {
                    tableParser.PushState(cell.CellPosition.Add(new ObjectSize(0, i)), new Styler(cell));
                    if(parser.TryParse(tableParser, type, out var result) && result != null)
                        parsed = true;
                    tableParser.PopState();
                }
                if(!parsed)
                    return i;
            }
            throw new EnumerableTooLongException(maxEnumerableLength);
        }
        private readonly ParserCollection parserCollection;
        private const int maxEnumerableLength = 200;
    }
}