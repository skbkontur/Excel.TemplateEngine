using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator
{
    public interface ITableNavigator
    {
        void PushState(ICellPosition newOrigin);
        void PushState();
        void PopState();
        void MoveToNextLayer();
        void MoveToNextColumn();
        TableNavigatorState CurrentState { get; }
    }
}