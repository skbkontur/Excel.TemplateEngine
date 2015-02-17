using System.Collections.Generic;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces
{
    public interface ITablePart
    {
        IEnumerable<IEnumerable<ICell>> Cells { get; }
    }
}