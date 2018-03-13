using System.Collections.Concurrent;
using System.Linq;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates
{
    public class TemplateCollection : ITemplateCollection
    {
        public TemplateCollection(ITable templateTable)
        {
            this.templateTable = templateTable;
            cache = new ConcurrentDictionary<string, RenderingTemplate>();
        }

        public RenderingTemplate GetTemplate(string templateName)
        {
            if(string.IsNullOrEmpty(templateName))
                return null;

            if(cache.TryGetValue(templateName, out var template))
                return template;

            AddNewTemplateIntoCache(templateName);

            return cache[templateName];
        }

        private void AddNewTemplateIntoCache(string templateName)
        {
            cache.TryAdd(templateName, null);

            var cell = SearchTemplateDescription(templateName);
            if(cell == null)
                return;

            if(!TemplateDescriptionHelper.TryExtractCoordinates(cell.StringValue, out var range))
                return;

            var newTemplate = BuildNewRenderingTemplate(range);

            if(newTemplate.IsValid())
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
            var templateNameWithPrefix = string.Format("Template:{0}:", templateName);
            return templateTable.SearchCellByText(templateNameWithPrefix).FirstOrDefault();
        }

        private readonly ITable templateTable;
        private readonly ConcurrentDictionary<string, RenderingTemplate> cache;
    }
}