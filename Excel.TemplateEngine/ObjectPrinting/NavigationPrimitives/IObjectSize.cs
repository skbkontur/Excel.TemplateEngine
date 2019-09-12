namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives
{
    public interface IObjectSize
    {
        IObjectSize Add(IObjectSize other);
        IObjectSize Subtract(IObjectSize other);
        int Width { get; }
        int Height { get; }
    }
}