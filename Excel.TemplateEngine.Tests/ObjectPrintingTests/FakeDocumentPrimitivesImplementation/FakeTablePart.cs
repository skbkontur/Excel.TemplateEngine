using System.Collections.Generic;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;

namespace SkbKontur.Excel.TemplateEngine.Tests.ObjectPrintingTests.FakeDocumentPrimitivesImplementation
{
    internal class FakeTablePart : ITablePart
    {
        public FakeTablePart(IReadOnlyList<IReadOnlyList<ICell>> cells)
        {
            Cells = cells;
        }

        public IReadOnlyList<IReadOnlyList<ICell>> Cells { get; }
    }
}