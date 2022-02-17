using System.Collections.Generic;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives
{
    public interface ITablePart //TODO: {birne} порвать дубликаты
    {
        IReadOnlyList<IReadOnlyList<ICell>> Cells { get; }
    }
}