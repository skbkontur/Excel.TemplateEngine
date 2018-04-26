using System;
using System.Collections.Generic;

using JetBrains.Annotations;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Interfaces
{
    public interface IExcelWorksheet
    {
        void SetPrinterSettings(ExcelPrinterSettings excelPrinterSettings);
        IExcelCell InsertCell(ExcelCellIndex cellIndex);
        IExcelRow CreateRow(int rowIndex);
        void MergeCells(ExcelCellIndex upperLeft, ExcelCellIndex lowerRight);
        void ResizeColumn(int columnIndex, double width);
        IEnumerable<IExcelCell> GetSortedCellsInRange(ExcelCellIndex upperLeft, ExcelCellIndex lowerRight);
        IExcelCell GetCell(ExcelCellIndex position);
        IEnumerable<IExcelCell> SearchCellsByText(string text);
        IEnumerable<IExcelRow> Rows { get; }
        IEnumerable<IExcelColumn> Columns { get; }
        IEnumerable<Tuple<ExcelCellIndex, ExcelCellIndex>> MergedCells { get; }
        IExcelDocument ExcelDocument { get; }

        [CanBeNull]
        IExcelCheckBoxControlInfo TryGetCheckBoxFormControlInfo([NotNull] string name);

        [CanBeNull]
        IExcelDropDownControlInfo TryGetDropDownFormControlInfo([NotNull] string name);

        void CopyFormControlsFrom([NotNull] IExcelWorksheet template);
        void CopyDataValidationsFrom([NotNull] IExcelWorksheet template);
        void CopyWorksheetExtensionListFrom([NotNull] IExcelWorksheet template);
    }
}