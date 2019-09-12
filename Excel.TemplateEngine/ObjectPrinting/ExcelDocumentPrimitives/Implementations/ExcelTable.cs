using System.Collections.Generic;
using System.Linq;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;
using SkbKontur.Excel.TemplateEngine.FileGenerating.Primitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives.Implementations
{
    public class ExcelTable : ITable
    {
        public ExcelTable([NotNull] IExcelWorksheet excelWorksheet)
        {
            internalTable = excelWorksheet;
        }

        public ICell GetCell(ICellPosition position)
        {
            var internalCell = internalTable.GetCell(new ExcelCellIndex(position.CellReference));
            return internalCell == null ? null : new ExcelCell(internalCell);
        }

        public ICell InsertCell(ICellPosition position)
        {
            return new ExcelCell(internalTable.InsertCell(new ExcelCellIndex(position.CellReference)));
        }

        public IEnumerable<ICell> SearchCellByText(string text)
        {
            return internalTable.SearchCellsByText(text).Select(cell => new ExcelCell(cell));
        }

        public ITablePart GetTablePart(IRectangle rectangle)
        {
            var excelReferenceToCell = internalTable
                                       .GetSortedCellsInRange(new ExcelCellIndex(rectangle.UpperLeft.CellReference),
                                                              new ExcelCellIndex(rectangle.LowerRight.CellReference))
                                       .Select(cell => new ExcelCell(cell))
                                       .ToDictionary(cell => cell.CellPosition.CellReference);

            var subTableSize = rectangle.Size;
            var subTable = JaggedArrayHelper.CreateJaggedArray<ICell[][]>(subTableSize.Height, subTableSize.Width);
            for (var x = rectangle.UpperLeft.ColumnIndex; x <= rectangle.LowerRight.ColumnIndex; ++x)
            {
                for (var y = rectangle.UpperLeft.RowIndex; y <= rectangle.LowerRight.RowIndex; ++y)
                {
                    if (!excelReferenceToCell.TryGetValue(new CellPosition(y, x).CellReference, out var cell))
                        cell = new ExcelCell(internalTable.InsertCell(new ExcelCellIndex(y, x)));

                    subTable[y - rectangle.UpperLeft.RowIndex][x - rectangle.UpperLeft.ColumnIndex] = cell;
                }
            }

            return new ExcelTablePart(subTable);
        }

        public IEnumerable<IColumn> Columns { get { return internalTable.Columns.Select(c => new ExcelColumn(c)); } }

        public void MergeCells(IRectangle rectangle)
        {
            internalTable.MergeCells(new ExcelCellIndex(rectangle.UpperLeft.CellReference),
                                     new ExcelCellIndex(rectangle.LowerRight.CellReference));
        }

        public void CopyFormControlsFrom([NotNull] ITable template)
        {
            internalTable.CopyFormControlsFrom(((ExcelTable)template).internalTable);
        }

        public void CopyDataValidationsFrom([NotNull] ITable template)
        {
            internalTable.CopyDataValidationsFrom(((ExcelTable)template).internalTable);
        }

        public void CopyWorksheetExtensionListFrom(ITable template)
        {
            internalTable.CopyWorksheetExtensionListFrom(((ExcelTable)template).internalTable);
        }

        public void CopyCommentsFrom([NotNull] ITable template)
        {
            internalTable.CopyComments(((ExcelTable)template).internalTable);
        }

        [CanBeNull]
        public IExcelCheckBoxControlInfo TryGetCheckBoxFormControl([NotNull] string name)
        {
            return internalTable.TryGetCheckBoxFormControlInfo(name);
        }

        [CanBeNull]
        public IExcelDropDownControlInfo TryGetDropDownFormControl([NotNull] string name)
        {
            return internalTable.TryGetDropDownFormControlInfo(name);
        }

        public void ResizeColumn(int columnIndex, double width)
        {
            internalTable.ResizeColumn(columnIndex, width);
        }

        public IEnumerable<IRectangle> MergedCells
        {
            get
            {
                return internalTable.MergedCells
                                    .Select(tuple => new Rectangle(new CellPosition(tuple.Item1),
                                                                   new CellPosition(tuple.Item2)));
            }
        }

        [NotNull]
        private readonly IExcelWorksheet internalTable;
    }
}