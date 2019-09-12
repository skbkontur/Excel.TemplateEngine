using SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations
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