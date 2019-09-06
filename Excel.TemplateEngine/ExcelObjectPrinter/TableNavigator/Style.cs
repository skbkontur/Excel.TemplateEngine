using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator
{
    public class Style : IStyle
    {
        public Style(ICell templateCell)
        {
            this.templateCell = templateCell;
        }

        public void ApplyTo(ICell cell)
        {
            cell.CopyStyle(templateCell);
        }

        private readonly ICell templateCell;
    }
}