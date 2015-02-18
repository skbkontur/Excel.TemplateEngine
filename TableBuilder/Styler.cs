using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public class Styler : IStyler
    {
        public Styler(ICell templateCell)
        {
            this.templateCell = templateCell;
        }

        public void ApplyStyle(ICell cell)
        {
            cell.CopyStyle(templateCell);
        }

        private readonly ICell templateCell;
    }
}