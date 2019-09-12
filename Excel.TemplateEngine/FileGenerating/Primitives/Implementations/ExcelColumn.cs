namespace Excel.TemplateEngine.FileGenerating.Primitives.Implementations
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