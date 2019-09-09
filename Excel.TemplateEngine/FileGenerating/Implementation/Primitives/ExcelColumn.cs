using Excel.TemplateEngine.FileGenerating.Interfaces;

namespace Excel.TemplateEngine.FileGenerating.Implementation.Primitives
{
    public class ExcelColumn : IExcelColumn
    {
        public ExcelColumn(double width, int index)
        {
            Width = width;
            Index = index;
        }

        public double Width { get; }
        public int Index { get; }
    }
}