using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;

namespace Excel.TemplateEngine.Tests.ObjectPrintingTests.FakeDocumentPrimitivesImplementation
{
    public class FakeColumn : IColumn
    {
        public int Index { get; set; }
        public double Width => 1.0;
    }
}