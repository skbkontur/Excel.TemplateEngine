using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation
{
    public class ExcelColumn : IColumn
    {
        public ExcelColumn(IExcelColumn internalColumn)
        {
            this.internalColumn = internalColumn;
        }

        public int Index { get { return internalColumn.Index; } }
        public double Width { get { return internalColumn.Width; } }

        private readonly IExcelColumn internalColumn;
    }
}