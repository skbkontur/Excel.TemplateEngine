using System.Collections.Generic;
using System.Linq;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates
{
    internal class RenderingTemplate
    {
        public IRectangle Range { get; set; }
        public ITablePart Content { get; set; }
        public IEnumerable<IRectangle> MergedCells { get; set; }
        public IEnumerable<IColumn> Columns { get; set; }
    }

    internal static class RenderingTemplateExtensions
    {
        public static bool IsValid(this RenderingTemplate renderingTemplate) //TODO: {birne} здесь нужно ещё проверить, что во всех клетках слитых областей только атомарные значения
        {
            return renderingTemplate.MergedCells.All(renderingTemplate.IsMergedCellInBounds);
        }

        private static bool IsMergedCellInBounds(this RenderingTemplate renderingTemplate, IRectangle mergedCellRelativePosition)
        {
            var origin = renderingTemplate.Range.UpperLeft;
            return renderingTemplate.Range.Contains(mergedCellRelativePosition.ToGlobalCoordinates(origin));
        }
    }
}