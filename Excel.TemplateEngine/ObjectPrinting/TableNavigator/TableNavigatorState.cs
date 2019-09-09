using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace Excel.TemplateEngine.ObjectPrinting.TableNavigator
{
    public class TableNavigatorState
    {
        public ICellPosition Origin { get; set; }
        public ICellPosition Cursor { get; set; }
        public int CurrentLayerStartRowIndex { get; set; }
        public int CurrentLayerHeight { get; set; }
        public int GlobalWidth { get; set; }
        public int GlobalHeight { get; set; }
    }
}