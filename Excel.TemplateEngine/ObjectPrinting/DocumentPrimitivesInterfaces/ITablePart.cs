using System.Collections.Generic;

namespace Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces
{
    public interface ITablePart //TODO: {birne} порвать дубликаты
    {
        IEnumerable<IEnumerable<ICell>> Cells { get; }
    }
}