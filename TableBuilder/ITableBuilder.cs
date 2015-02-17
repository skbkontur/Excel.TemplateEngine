using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public interface ITableBuilder
    {
        ITableBuilder RenderAtomicValue(string value);
        ITableBuilder RenderAtomicValue(int value);
        ITableBuilder RenderAtomicValue(double value);
        ITableBuilder RenderAtomicValue(decimal value);
        ITableBuilder PushState(ICellPosition newOrigin);
        ITableBuilder PushState();
        ITableBuilder PopState();
        ITableBuilder MoveToNextLayer();
    }
}