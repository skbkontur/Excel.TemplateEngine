using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public interface ITableBuilder
    {
        ITableBuilder RenderAtomicValue(string value);
        ITableBuilder RenderAtomicValue(int value);
        ITableBuilder RenderAtomicValue(double value);
        ITableBuilder RenderAtomicValue(decimal value);
        ITableBuilder PushState(ICellPosition newOrigin, IStyler styler);
        ITableBuilder PushState(IStyler styler);
        ITableBuilder PushState();
        ITableBuilder PopState();
        ITableBuilder MoveToNextLayer();
        ITableBuilder MoveToNextColumn();
        ITableBuilder ResizeColumn(int columnIndex, double width);
        ITableBuilder SetCurrentStyle();
        TableBuilderState CurrentState { get; }
        void MergeCells(IRectangle rectangle);
    }
}