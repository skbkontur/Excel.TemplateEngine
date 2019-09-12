using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser
{
    public interface ITableParser
    {
        bool TryParseAtomicValue(out string result);
        bool TryParseAtomicValue(out int result);
        bool TryParseAtomicValue(out double result);
        bool TryParseAtomicValue(out decimal result);
        bool TryParseAtomicValue(out long result);
        bool TryParseAtomicValue(out int? result);
        bool TryParseAtomicValue(out double? result);
        bool TryParseAtomicValue(out decimal? result);
        bool TryParseAtomicValue(out long? result);
        bool TryParseCheckBoxValue(string name, out bool result);
        bool TryParseDropDownValue(string name, out string result);
        ITableParser PushState(ICellPosition newOrigin);
        ITableParser PushState();
        ITableParser PopState();
        ITableParser MoveToNextLayer();
        ITableParser MoveToNextColumn();
        TableNavigatorState CurrentState { get; }
    }
}