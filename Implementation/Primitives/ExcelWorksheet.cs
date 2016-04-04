using System;
using System.Collections.Generic;
using System.Linq;

using C5;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.DataTypes;
using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.Objects;

using Tuple = System.Tuple;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    public class ExcelWorksheet : IExcelWorksheet
    {
        public ExcelWorksheet(WorksheetPart worksheetPart, IExcelDocumentStyle documentStyle, IExcelSharedStrings excelSharedStrings)
        {
            worksheet = worksheetPart.Worksheet;
            this.documentStyle = documentStyle;
            this.excelSharedStrings = excelSharedStrings;
            rowsCache = new TreeDictionary<uint, Row>();
            var sheetData = worksheet.GetFirstChild<SheetData>();
            if(sheetData != null)
                rowsCache.AddAll(sheetData.Elements<Row>().Select(x => new C5.KeyValuePair<uint, Row>(x.RowIndex, x)));
        }

        public void SetPrinterSettings(ExcelPrinterSettings excelPrinterSettings)
        {
            if(excelPrinterSettings.PageMargins != null)
            {
                var pageMargins = worksheet.Elements<PageMargins>().FirstOrDefault() ?? new PageMargins();
                pageMargins.Left = excelPrinterSettings.PageMargins.Left;
                pageMargins.Right = excelPrinterSettings.PageMargins.Right;
                pageMargins.Top = excelPrinterSettings.PageMargins.Top;
                pageMargins.Bottom = excelPrinterSettings.PageMargins.Bottom;
                pageMargins.Header = excelPrinterSettings.PageMargins.Header;
                pageMargins.Footer = excelPrinterSettings.PageMargins.Footer;

                if(!worksheet.Elements<PageMargins>().Any())
                    worksheet.AppendChild(pageMargins);
            }

            var pageSetup = worksheet.Elements<PageSetup>().FirstOrDefault() ?? new PageSetup();
            pageSetup.Orientation = excelPrinterSettings.PrintingOrientation == ExcelPrintingOrientation.Landscape ? OrientationValues.Landscape : OrientationValues.Portrait;

            if(!worksheet.Elements<PageSetup>().Any())
                worksheet.AppendChild(pageSetup);
        }

        public void MergeCells(ExcelCellIndex upperLeft, ExcelCellIndex lowerRight)
        {
            var mergeCells = worksheet.GetFirstChild<MergeCells>() ?? CreateMergeCellsWorksheetPart();
            mergeCells.AppendChild(new MergeCell
                {
                    Reference = string.Format("{0}:{1}", upperLeft.CellReference, lowerRight.CellReference)
                });
        }

        public void ResizeColumn(int columnIndex, double width)
        {
            var columns = worksheet.GetFirstChild<Columns>() ?? CreateColumns();
            while(columns.ChildElements.Count < columnIndex)
            {
                columns.AppendChild(new Column
                    {
                        Min = (uint)columns.ChildElements.Count + 1,
                        Max = (uint)columns.ChildElements.Count + 1,
                        BestFit = true,
                        CustomWidth = true,
                        Width = 8.43
                    });
            }
            var column = (Column)columns.ChildElements.Skip(columnIndex - 1).First();
            column.Width = width;
            if(Math.Abs(width - 0) < 1e-9)
                column.Hidden = true;
        }

        public IEnumerable<IExcelCell> GetSortedCellsInRange(ExcelCellIndex upperLeft, ExcelCellIndex lowerRight)
        {
            return rowsCache.RangeFromTo((uint)upperLeft.RowIndex, (uint)lowerRight.RowIndex + 1)
                            .Select(x => x.Value)
                            .SelectMany(row => row.Elements<Cell>()
                                                  .Where(cell =>
                                                      {
                                                          var columnIndex = new ExcelCellIndex(cell.CellReference).ColumnIndex;
                                                          return columnIndex >= upperLeft.ColumnIndex && columnIndex <= lowerRight.ColumnIndex;
                                                      }))
                            .OrderBy(cell =>
                                {
                                    var cellIndex = new ExcelCellIndex(cell.CellReference);
                                    return (cellIndex.RowIndex - upperLeft.RowIndex) * (lowerRight.ColumnIndex - upperLeft.ColumnIndex) + cellIndex.ColumnIndex;
                                })
                            .Select(cell => new ExcelCell(cell, documentStyle, excelSharedStrings));
        }

        public IExcelCell GetCell(ExcelCellIndex position)
        {
            return GetSortedCellsInRange(position, position).FirstOrDefault();
        }

        public IEnumerable<IExcelCell> SearchCellsByText(string text)
        {
            return rowsCache.Select(x => x.Value)
                            .SelectMany(row => row.Elements<Cell>())
                            .Select(internalCell => new ExcelCell(internalCell, documentStyle, excelSharedStrings))
                            .Where(cell => cell.GetStringValue().Return(str => str.Contains(text), false));
        }

        public IEnumerable<IExcelRow> Rows { get { return rowsCache.Select(x => new ExcelRow(x.Value, documentStyle, excelSharedStrings)); } }

        public IEnumerable<IExcelColumn> Columns
        {
            get
            {
                if(worksheet.GetFirstChild<Columns>() == null)
                    return Enumerable.Empty<IExcelColumn>();
                return worksheet.GetFirstChild<Columns>().ChildElements.OfType<Column>().SelectMany(x =>
                    {
                        var list = new List<IExcelColumn>();
                        for(var index = x.Min; index <= x.Max; ++index)
                            list.Add(new ExcelColumn(x.Width, (int)index.Value));
                        return list.ToArray();
                    });
            }
        }

        public IEnumerable<Tuple<ExcelCellIndex, ExcelCellIndex>> MergedCells
        {
            get
            {
                return worksheet.With(w => w.GetFirstChild<MergeCells>())
                                .Return(w => w.Select(x => (MergeCell)x), Enumerable.Empty<MergeCell>())
                                .Select(mergeCell => mergeCell.Reference.Value.Split(':').ToArray())
                                .Select(references => Tuple.Create(new ExcelCellIndex(references[0]), new ExcelCellIndex(references[1])));
            }
        }

        public IExcelCell InsertCell(ExcelCellIndex cellIndex)
        {
            var newRow = CreateRow(cellIndex.RowIndex);
            return newRow.CreateCell(cellIndex.ColumnIndex);
        }

        public IExcelRow CreateRow(int rowIndex)
        {
            var unsignedRowIndex = (uint)rowIndex;
            var sheetData = worksheet.GetFirstChild<SheetData>();
            C5.KeyValuePair<uint, Row> successor;
            Row refRow = null;
            var newRow = new Row
                {
                    RowIndex = new UInt32Value((uint)rowIndex),
                };
            if(rowsCache.TryWeakSuccessor(unsignedRowIndex, out successor))
            {
                if(successor.Key == unsignedRowIndex)
                    return new ExcelRow(successor.Value, documentStyle, excelSharedStrings);
                refRow = successor.Value;
            }
            sheetData.InsertBefore(newRow, refRow);
            rowsCache.Add(unsignedRowIndex, newRow);
            return new ExcelRow(newRow, documentStyle, excelSharedStrings);
        }

        private MergeCells CreateMergeCellsWorksheetPart()
        {
            // Имеет принципиальное значение, куда именно вставлять элемент MergeCells
            // см. http://msdn.microsoft.com/en-us/library/office/cc880096(v=office.15).aspx

            var mergeCells = new MergeCells();
            if(worksheet.Elements<CustomSheetView>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<CustomSheetView>().First());
            else if(worksheet.Elements<DataConsolidate>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<DataConsolidate>().First());
            else if(worksheet.Elements<SortState>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SortState>().First());
            else if(worksheet.Elements<AutoFilter>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<AutoFilter>().First());
            else if(worksheet.Elements<Scenarios>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<Scenarios>().First());
            else if(worksheet.Elements<ProtectedRanges>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<ProtectedRanges>().First());
            else if(worksheet.Elements<SheetProtection>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetProtection>().First());
            else if(worksheet.Elements<SheetCalculationProperties>().Any())
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetCalculationProperties>().First());
            else
                worksheet.InsertAfter(mergeCells, worksheet.Elements<SheetData>().First());
            return mergeCells;
        }

        private Columns CreateColumns()
        {
            var columns = new Columns();
            if(worksheet.Elements<SheetFormatProperties>().Any())
                worksheet.InsertAfter(columns, worksheet.Elements<SheetFormatProperties>().First());
            else if(worksheet.Elements<SheetViews>().Any())
                worksheet.InsertAfter(columns, worksheet.Elements<SheetViews>().First());
            else if(worksheet.Elements<Dimensions>().Any())
                worksheet.InsertAfter(columns, worksheet.Elements<Dimensions>().First());
            else if(worksheet.Elements<SheetProperties>().Any())
                worksheet.InsertAfter(columns, worksheet.Elements<SheetProperties>().First());
            else
                worksheet.InsertAt(columns, 0);
            return columns;
        }

        private readonly IExcelDocumentStyle documentStyle;
        private readonly IExcelSharedStrings excelSharedStrings;
        private readonly Worksheet worksheet;
        private readonly TreeDictionary<uint, Row> rowsCache;
    }
}