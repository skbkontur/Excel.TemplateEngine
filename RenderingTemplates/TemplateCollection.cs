using System;
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

            var cell = SearchTemplateDescription(templateName);
            if(cell == null)
                return null;

            Tuple<ICellPosition, ICellPosition> range;
            if(TemplateDescriptionHelper.Instance.TryExtractCoordinates(cell.StringValue, out range))
            {
                cache.Add(templateName, new RenderingTemplate
                    {
                        Content = templateTable.GetTablePart(range.Item1, range.Item2)
                    });
                return cache[templateName];
            }

            return null;
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