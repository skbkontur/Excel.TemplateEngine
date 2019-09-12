using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator
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