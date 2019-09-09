using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;

namespace Excel.TemplateEngine.ObjectPrinting.TableNavigator
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