using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    public class SimpleCell
    {
        public SimpleCell([NotNull] ICellPosition cellPosition, [CanBeNull] string cellValue)
        {
            CellPosition = cellPosition;
            CellValue = cellValue;
        }

        [NotNull]
        public ICellPosition CellPosition { get; }

        [CanBeNull]
        public string CellValue { get; }
    }
}