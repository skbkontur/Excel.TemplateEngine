using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.FakeDocumentPrimitivesImplementation
{
    public class FakeColumn : IColumn
    {
        public int Index { get; set; }
        public double Width { get { return 1.0; } }
    }
}