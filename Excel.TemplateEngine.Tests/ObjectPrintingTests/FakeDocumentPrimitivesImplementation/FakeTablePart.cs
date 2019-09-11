using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests.FakeDocumentPrimitivesImplementation
{
    public class FakeTablePart : ITablePart
    {
        public FakeTablePart(IEnumerable<IEnumerable<ICell>> cells)
        {
            Cells = cells;
        }

        public IEnumerable<IEnumerable<ICell>> Cells { get; }
    }
}