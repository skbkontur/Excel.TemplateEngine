using System.Collections.Generic;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives
{
    public interface ITablePart //TODO: {birne} порвать дубликаты
    {
        IEnumerable<IEnumerable<ICell>> Cells { get; }
    }
}