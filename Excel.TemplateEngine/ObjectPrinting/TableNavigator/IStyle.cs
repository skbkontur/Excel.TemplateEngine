using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;

namespace Excel.TemplateEngine.ObjectPrinting.TableNavigator
{
    public interface IStyle
    {
        void ApplyTo(ICell cell);
    }
}