using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator
{
    public interface ITableNavigator
    {
        void PushState(ICellPosition newOrigin, IStyler styler);
        void PushState(IStyler styler);
        void PushState();
        void PopState();
        void MoveToNextLayer();
        void MoveToNextColumn();
        void SetCurrentStyle();
        TableNavigatorState CurrentState { get; }
        ITable Target { get; }
    }
}