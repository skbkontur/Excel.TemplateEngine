using System.Collections.Generic;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates
{
    public class RenderingTemplate
    {
        public ITablePart Content { get; set; }
        public IEnumerable<IRectangle> MergedCells { get; set; }
    }
}