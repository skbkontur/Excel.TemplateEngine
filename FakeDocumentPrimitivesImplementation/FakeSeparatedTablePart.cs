using System.Collections.Generic;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation
{
    public class FakeSeparatedTablePart : ITablePart
    {
        public FakeSeparatedTablePart(IEnumerable<IEnumerable<ICell>> cells)
        {
            Cells = cells;
        }

        public IEnumerable<IEnumerable<ICell>> Cells { get; private set; }
    }
}