using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using Excel.TemplateEngine.ObjectPrinting.TableNavigator;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.TableBuilder
{
    public interface ITableBuilder
    {
        ITableBuilder RenderAtomicValue(string value);
        ITableBuilder RenderAtomicValue(int value);
        ITableBuilder RenderAtomicValue(double value);
        ITableBuilder RenderAtomicValue(decimal value);

        [NotNull]
        ITableBuilder RenderCheckBoxValue([NotNull] string name, bool value);

        [NotNull]
        ITableBuilder RenderDropDownValue([NotNull] string name, [CanBeNull] string value);

        ITableBuilder PushState(ICellPosition newOrigin, IStyle style);
        ITableBuilder PushState(IStyle style);
        ITableBuilder PushState();
        ITableBuilder PopState();
        ITableBuilder MoveToNextLayer();
        ITableBuilder MoveToNextColumn();
        ITableBuilder ExpandColumn(int relativeColumnIndex, double width);
        ITableBuilder SetCurrentStyle();
        ITableBuilder MergeCells(IRectangle rectangle);
        ITableBuilder CopyFormControlsFrom([NotNull] ITable template);
        ITableBuilder CopyDataValidationsFrom([NotNull] ITable template);
        ITableBuilder CopyWorksheetExtensionListFrom([NotNull] ITable template);
        ITableBuilder CopyCommentsFrom([NotNull] ITable template);
    }
}