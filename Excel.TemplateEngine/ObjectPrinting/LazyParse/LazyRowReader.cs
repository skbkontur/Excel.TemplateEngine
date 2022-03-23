using System;
using System.Collections.Generic;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    /// <summary>
    /// Reads current or following cell. Can't read previous ones.
    /// </summary>
    public class LazyRowReader
    {
        public LazyRowReader([NotNull] Row row, [NotNull] IReadOnlyList<string> sharedStrings)
        {
            RowIndex = (int)row.RowIndex!.Value;
            this.sharedStrings = sharedStrings;
            reader = OpenXmlReader.Create(row);
        }

        [NotNull]
        private SimpleCell LoadCurrentCell()
        {
            var cell = (Cell)reader.LoadCurrentElement();
            return ToSimpleCell(cell!);
        }

        [CanBeNull]
        public SimpleCell TryReadCell([NotNull] ICellPosition cellPosition)
        {
            if (cellPosition.RowIndex != RowIndex)
                throw new ArgumentException($"Incorrect cell reference. Target cell reference: {cellPosition}. Current row index: {RowIndex}.");

            if (currentCell != null)
            {
                if (cellPosition.ColumnIndex < currentCell.CellPosition.ColumnIndex)
                    return null;

                if (cellPosition.ColumnIndex == currentCell.CellPosition.ColumnIndex)
                    return currentCell;
            }

            while (reader.Read())
            {
                if (reader.ElementType != typeof(Cell))
                    continue;

                currentCell = LoadCurrentCell();

                if (cellPosition.ColumnIndex > currentCell!.CellPosition.ColumnIndex)
                    continue;

                if (cellPosition.ColumnIndex < currentCell.CellPosition.ColumnIndex)
                    return null;

                if (cellPosition.ColumnIndex == currentCell.CellPosition.ColumnIndex)
                    return currentCell;
            }

            return null;
        }

        [NotNull]
        private SimpleCell ToSimpleCell([NotNull] Cell cell)
        {
            var cellIndex = new CellPosition(cell.CellReference);
            var cellString = cell.InnerText;
            if (cell.DataType?.Value == CellValues.SharedString)
            {
                var i = int.Parse(cell.InnerText);
                cellString = sharedStrings[i];
            }

            return new SimpleCell(cellIndex, cellString);
        }

        public readonly int RowIndex;

        [NotNull]
        private readonly IReadOnlyList<string> sharedStrings;

        [NotNull]
        private readonly OpenXmlReader reader;

        [CanBeNull]
        private SimpleCell currentCell;
    }
}