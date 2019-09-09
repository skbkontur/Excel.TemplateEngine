using System.Collections.Generic;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces
{
    public interface ITablePart //TODO: {birne} порвать дубликаты
    {
        IEnumerable<IEnumerable<ICell>> Cells { get; }
    }
}