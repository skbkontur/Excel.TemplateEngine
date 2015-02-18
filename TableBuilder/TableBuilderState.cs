using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public class TableBuilderState
    {
        public ICellPosition Origin { get; set; }
        public ICellPosition Cursor { get; set; }
        public int CurrentLayerStartRowIndex { get; set; }
        public int CurrentLayerHeight { get; set; }
        public int GlobalWidth { get; set; }
        public int GlobalHeight { get; set; }
        public IStyler Styler { get; set; }
    }
}