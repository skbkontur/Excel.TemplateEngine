using System.Collections.Generic;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;

namespace Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitivesImplementation
{
    public class ExcelTablePart : ITablePart
    {
        public ExcelTablePart(IEnumerable<IEnumerable<ICell>> cells)
        {
            Cells = cells;
        }

        public IEnumerable<IEnumerable<ICell>> Cells { get; }
    }
}