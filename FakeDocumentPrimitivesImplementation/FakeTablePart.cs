using System.Collections.Generic;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation
{
    public class FakeTablePart : ITablePart
    {
        public FakeTablePart(IEnumerable<IEnumerable<ICell>> cells)
        {
            Cells = cells;
        }

        public IEnumerable<IEnumerable<ICell>> Cells { get; private set; }
    }
}