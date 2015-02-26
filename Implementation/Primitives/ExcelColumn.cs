using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelColumn : IExcelColumn
    {
        public ExcelColumn(double width, int index)
        {
            Width = width;
            Index = index;
        }

        public double Width { get; private set; }
        public int Index { get; private set; }
    }
}