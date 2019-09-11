using Excel.TemplateEngine.FileGenerating.Interfaces;
using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;

namespace Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitivesImplementation
{
    public class ExcelColumn : IColumn
    {
        public ExcelColumn(IExcelColumn internalColumn)
        {
            this.internalColumn = internalColumn;
        }

        public int Index => internalColumn.Index;
        public double Width => internalColumn.Width;

        private readonly IExcelColumn internalColumn;
    }
}