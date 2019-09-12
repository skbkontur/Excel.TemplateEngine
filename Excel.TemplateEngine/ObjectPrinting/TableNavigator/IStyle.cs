using Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;

namespace Excel.TemplateEngine.ObjectPrinting.TableNavigator
{
    public interface IStyle
    {
        void ApplyTo(ICell cell);
    }
}