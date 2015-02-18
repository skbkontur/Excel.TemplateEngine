using System.Linq;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.PostBuildActions
{
    public class ColumnResizer : IColumnResizer
    {
        public ColumnResizer(ITable templateTable)
        {
            this.templateTable = templateTable;
        }

        public void ResizeColumns(ITableBuilder tableBuilder)
        {
            foreach(var column in templateTable.Columns.ToArray())
                tableBuilder.ResizeColumn(column.Index, column.Width);
        }

        private readonly ITable templateTable;
    }
}