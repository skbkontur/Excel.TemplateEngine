using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;

namespace Excel.TemplateEngine.ObjectPrinting.FakeDocumentPrimitivesImplementation
{
    public class FakeColumn : IColumn
    {
        public int Index { get; set; }
        public double Width { get { return 1.0; } }
    }
}