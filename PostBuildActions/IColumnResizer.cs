using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.PostBuildActions
{
    public interface IColumnResizer
    {
        void ResizeColumns(ITableBuilder tableBuilder);
    }
}