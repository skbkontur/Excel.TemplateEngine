using System.Collections.Generic;
using System.Linq;

using Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using Excel.TemplateEngine.ObjectPrinting.Helpers;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;

namespace Excel.TemplateEngine.ObjectPrinting.RenderingTemplates
{
    internal class TemplateCollection : ITemplateCollection
    {
        public TemplateCollection(ITable templateTable)
        {
            this.templateTable = templateTable;
            cache = new Dictionary<string, RenderingTemplate>();
        }

        public RenderingTemplate GetTemplate(string templateName)
        {
            if (string.IsNullOrEmpty(templateName))
                return null;

            if (cache.TryGetValue(templateName, out var template))
                return template;

            AddNewTemplateIntoCache(templateName);

            return cache[templateName];
        }

        private void AddNewTemplateIntoCache(string templateName)
        {
            cache.Add(templateName, null);

            var cell = SearchTemplateDescription(templateName);
            if (cell == null)
                return;

            if (!TemplateDescriptionHelper.TryExtractCoordinates(cell.StringValue, out var range))
                return;

            var newTemplate = BuildNewRenderingTemplate(range);

            if (newTemplate.IsValid())
                cache[templateName] = newTemplate;
        }

        private RenderingTemplate BuildNewRenderingTemplate(IRectangle range)
        {
            return new RenderingTemplate
                {
                    Range = range,
                    Content = templateTable.GetTablePart(range),
                    MergedCells = templateTable.MergedCells
                                               .Where(rect => rect.Intersects(range))
                                               .Select(rect => rect.ToRelativeCoordinates(range.UpperLeft)),
                    Columns = templateTable.Columns
                                           .Where(column => column.Index >= range.UpperLeft.ColumnIndex &&
                                                            column.Index <= range.LowerRight.ColumnIndex)
                };
        }

        private ICell SearchTemplateDescription(string templateName)
        {
            var templateNameWithPrefix = $"Template:{templateName}:";
            return templateTable.SearchCellByText(templateNameWithPrefix).FirstOrDefault();
        }

        private readonly ITable templateTable;
        private readonly Dictionary<string, RenderingTemplate> cache;
    }
}