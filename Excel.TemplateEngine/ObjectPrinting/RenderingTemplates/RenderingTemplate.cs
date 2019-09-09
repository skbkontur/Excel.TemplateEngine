using System.Collections.Generic;
using System.Linq;

using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace Excel.TemplateEngine.ObjectPrinting.RenderingTemplates
{
    public class RenderingTemplate
    {
        public IRectangle Range { get; set; }
        public ITablePart Content { get; set; }
        public IEnumerable<IRectangle> MergedCells { get; set; }
        public IEnumerable<IColumn> Columns { get; set; }
    }

    public static class RenderingTemplateExtensions
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