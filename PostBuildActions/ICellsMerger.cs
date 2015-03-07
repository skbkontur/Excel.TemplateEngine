using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.PostBuildActions
{
    public interface ICellsMerger
    {
        void MergeCells(ITableBuilder tableBuilder);
    }
}