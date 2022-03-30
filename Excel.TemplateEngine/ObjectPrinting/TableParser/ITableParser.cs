using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser
{
    public interface ITableParser
    {
        bool TryParseCheckBoxValue(string name, out bool result);
        bool TryParseDropDownValue(string name, out string result);

        [NotNull]
        ITableParser PushState(ICellPosition newOrigin);

        [NotNull]
        ITableParser PushState();

        [NotNull]
        ITableParser PopState();

        [NotNull]
        ITableParser MoveToNextLayer();

        [NotNull]
        ITableParser MoveToNextColumn();

        [CanBeNull]
        public string GetCurrentCellText();

        TableNavigatorState CurrentState { get; }
    }
}