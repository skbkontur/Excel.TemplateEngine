using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator
{
    public interface IStyle
    {
        void ApplyTo(ICell cell);
    }
}