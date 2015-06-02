using System.Collections.Generic;
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
            cache = new Dictionary<string, RenderingTemplate>();
        }

        public RenderingTemplate GetTemplate(string templateName)
        {
            if(string.IsNullOrEmpty(templateName))
                return null;

            RenderingTemplate template;
            if(cache.TryGetValue(templateName, out template))
                return template;

            AddNewTemplateIntoCache(templateName);

            return cache[templateName];
        }

        private void AddNewTemplateIntoCache(string templateName)
        {
            cache.Add(templateName, null);

            var cell = SearchTemplateDescription(templateName);
            if(cell == null)
                return;

            IRectangle range;
            if(!TemplateDescriptionHelper.Instance.TryExtractCoordinates(cell.StringValue, out range))
                return;

            var newTemplate = BuildNewRenderinGTemplate(range);

            if(newTemplate.IsValid())
                cache[templateName] = newTemplate;
        }

        private RenderingTemplate BuildNewRenderinGTemplate(IRectangle range)
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
        private readonly Dictionary<string, RenderingTemplate> cache;
    }
}