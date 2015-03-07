using System;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.PostBuildActions
{
    public class CellsMerger : ICellsMerger
    {
        public CellsMerger(ITable templateTable)
        {
            this.templateTable = templateTable;
        }

        public void MergeCells(ITableBuilder tableBuilder)
        {
            foreach(var cell in templateTable.SearchCellByText("MergeCells:"))
            {
                Tuple<ICellPosition, ICellPosition> range;
                if(TemplateDescriptionHelper.Instance.TryExtractCoordinates(cell.StringValue, out range))
                    tableBuilder.MergeCells(range.Item1, range.Item2);
            }
        }

        private readonly ITable templateTable;
    }
}