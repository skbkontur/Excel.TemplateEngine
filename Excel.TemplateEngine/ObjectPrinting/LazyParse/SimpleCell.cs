using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.LazyParse
{
    public class SimpleCell
    {
        public SimpleCell([NotNull] CellPosition cellPosition, [NotNull] string cellValue)
        {
            CellPosition = cellPosition;
            CellValue = cellValue;
        }

        [NotNull]
        public CellPosition CellPosition { get; }

        [NotNull]
        public string CellValue { get; }
    }
}