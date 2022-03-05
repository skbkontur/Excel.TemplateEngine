using System;
using System.Collections;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.FileGenerating.DataTypes;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ParseCollection.Parsers.Implementations
{
    class LazyRowReader : IEnumerable<Cell>
    {
        public LazyRowReader([NotNull] Row row)
        {
            this.row = row;
            reader = OpenXmlReader.Create(row);
        }

        [NotNull]
        public Cell GetCurrentCell()
        {
            if (reader.ElementType != typeof(Cell))
                throw new Exception($"Current element is not {nameof(Cell)}");

            return ((Cell)reader.LoadCurrentElement())!;
        }

        [NotNull]
        public Cell GetNextCell([CanBeNull] ICellPosition targetCellPosition = null)
        {
            if (targetCellPosition != null)
            {
                if (targetCellPosition.RowIndex != row.RowIndex!.Value)
                    throw new ArgumentException($"Incorrect cell reference. Target cell reference: {targetCellPosition}. Current row index: {row.RowIndex.Value}.");
            }

            while (reader.Read())
            {
                if (reader.ElementType != typeof(Cell))
                    continue;

                var cell = (Cell)reader.LoadCurrentElement();

                if (targetCellPosition != null)
                {
                    var currentCellIndex = new ExcelCellIndex(cell!.CellReference);
                    if (targetCellPosition!.ColumnIndex < currentCellIndex.ColumnIndex)
                        throw new ArgumentException("Cell column index is less than current cell column index. Can't read previous cells.");

                    if (targetCellPosition!.ColumnIndex > currentCellIndex.ColumnIndex)
                        continue;
                }

                return cell!;
            }
            throw new IndexOutOfRangeException(targetCellPosition?.CellReference);
        }

        public IEnumerator<Cell> GetEnumerator()
        {
            while (reader.Read())
            {
                if (reader.ElementType != typeof(Cell))
                    continue;

                yield return (Cell)reader.LoadCurrentElement();
            }
        }

        IEnumerator IEnumerable.GetEnumerator()
        {
            return GetEnumerator();
        }

        private readonly Row row;
        private readonly OpenXmlReader reader;
    }
}