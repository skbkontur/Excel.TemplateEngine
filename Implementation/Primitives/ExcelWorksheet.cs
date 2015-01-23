using System;
using System.Collections.Generic;
using System.Linq;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Caches;
using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;

namespace SKBKontur.Catalogue.ExcelFileGenerator.Implementation.Primitives
{
    internal class ExcelWorksheet : IExcelWorksheet
    {
        public ExcelWorksheet(WorksheetPart worksheetPart, IExcelDocumentStyle documentStyle, IExcelSharedStrings excelSharedStrings, IExcelDocumentMeta document)
        {
            worksheet = worksheetPart.Worksheet;
            this.documentStyle = documentStyle;
            this.excelSharedStrings = excelSharedStrings;
            this.document = document;
        }

        public void MergeCells(int fromRow, int fromCol, int toRow, int toCol)
        {
            var mergeCells = worksheet.GetFirstChild<MergeCells>() ?? CreateMergeCellsWorksheetPart();
            mergeCells.AppendChild(new MergeCell
                {
                    Reference = string.Format("{0}:{1}", IndexHelpers.ToCellName(fromRow, fromCol), IndexHelpers.ToCellName(toRow, toCol))
                });
        }

        public void CreateAutofilter(int fromRow, int fromCol, int toRow, int toCol)
        {
            var autofilter = worksheet.GetFirstChild<AutoFilter>() ?? CreateAutofilter();
            autofilter.Reference = string.Format("{0}:{1}", IndexHelpers.ToCellName(fromRow, fromCol), IndexHelpers.ToCellName(toRow, toCol));
        }

        public void CreateHyperlink(int row, int col, int toWorksheet, int toRow, int toCol)
        {
            var hyperlinks = worksheet.GetFirstChild<Hyperlinks>() ?? CreateHyperlinksWorksheetPart();
            var worksheetName = document.GetWorksheetName(toWorksheet);
            hyperlinks.AppendChild(new Hyperlink
                {
                    Reference = IndexHelpers.ToCellName(row, col),
                    Location = string.Format("{0}!{1}", worksheetName, IndexHelpers.ToCellName(toRow, toCol))
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
                        CustomWidth = true
                    });
            }
            var column = (Column)columns.ChildElements.Skip(columnIndex - 1).First();
            column.Width = width;
            if(Math.Abs(width - 0) < 1e-9)
                column.Hidden = true;
        }

        public IEnumerable<IExcelCell> GetSortedCellsInRange(int fromRow, int fromColumn, int toRow, int toColumn)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>()
                            .Where(row => row.RowIndex - 1 >= fromRow && row.RowIndex - 1 <= toRow)
                            .SelectMany(row => row.Elements<Cell>()
                                                  .Where(cell =>
                                                      {
                                                          var columnIndex = new ExcelCellIndex(cell.CellReference).ColumnIndex;
                                                          return columnIndex >= upperLeft.ColumnIndex && columnIndex <= lowerRight.ColumnIndex;
                                                      }))
                            .OrderBy(cell => (IndexHelpers.GetRowIndex(cell.CellReference) - fromRow) * (toColumn - fromColumn) + IndexHelpers.GetColumnIndex(cell.CellReference))
                            .Select(cell => new ExcelCell(cell, documentStyle, excelSharedStrings));
        }

        public IEnumerable<IExcelCell> GetSortedCellsInRange(string upperLeft, string lowerRight)
        {
            return GetSortedCellsInRange(IndexHelpers.GetRowIndex(upperLeft),
                                         IndexHelpers.GetColumnIndex(upperLeft),
                                         IndexHelpers.GetRowIndex(lowerRight),
                                         IndexHelpers.GetColumnIndex(lowerRight));
        }

        public IEnumerable<IExcelRow> Rows { get { return worksheet.GetFirstChild<SheetData>().ChildElements.OfType<Row>().Select(x => new ExcelRow(x, documentStyle, excelSharedStrings)); } }

        public IExcelRow CreateRow(int rowIndex)
        {
            var row = new Row
                {
                    RowIndex = new UInt32Value((uint)rowIndex)
                };
            worksheet.GetFirstChild<SheetData>().AppendChild(row);
            return new ExcelRow(row, documentStyle, excelSharedStrings);
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

        private Hyperlinks CreateHyperlinksWorksheetPart()
        {
            var hyperlinks = new Hyperlinks();
            if(worksheet.Elements<DataValidations>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<DataValidations>().First());
            else if(worksheet.Elements<ConditionalFormatting>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<ConditionalFormatting>().First());
            else if(worksheet.Elements<PhoneticProperties>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<PhoneticProperties>().First());
            else if(worksheet.Elements<MergeCells>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<MergeCells>().First());
            else if(worksheet.Elements<CustomSheetView>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<CustomSheetView>().First());
            else if(worksheet.Elements<DataConsolidate>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<DataConsolidate>().First());
            else if(worksheet.Elements<SortState>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SortState>().First());
            else if(worksheet.Elements<AutoFilter>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<AutoFilter>().First());
            else if(worksheet.Elements<Scenarios>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<Scenarios>().First());
            else if(worksheet.Elements<ProtectedRanges>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<ProtectedRanges>().First());
            else if(worksheet.Elements<SheetProtection>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SheetProtection>().First());
            else if(worksheet.Elements<SheetCalculationProperties>().Any())
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SheetCalculationProperties>().First());
            else
                worksheet.InsertAfter(hyperlinks, worksheet.Elements<SheetData>().First());
            return hyperlinks;
        }

        private AutoFilter CreateAutofilter()
        {
            var autofilter = new AutoFilter();
            if(worksheet.Elements<Scenarios>().Any())
                worksheet.InsertAfter(autofilter, worksheet.Elements<Scenarios>().First());
            else if(worksheet.Elements<ProtectedRanges>().Any())
                worksheet.InsertAfter(autofilter, worksheet.Elements<ProtectedRanges>().First());
            else if(worksheet.Elements<SheetProtection>().Any())
                worksheet.InsertAfter(autofilter, worksheet.Elements<SheetProtection>().First());
            else if(worksheet.Elements<SheetCalculationProperties>().Any())
                worksheet.InsertAfter(autofilter, worksheet.Elements<SheetCalculationProperties>().First());
            else
                worksheet.InsertAfter(autofilter, worksheet.Elements<SheetData>().First());
            return autofilter;
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
        private readonly IExcelDocumentMeta document;
    }
}