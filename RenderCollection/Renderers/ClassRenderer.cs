using System.Linq;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.Exceptions;
using SKBKontur.Catalogue.ExcelObjectPrinter.Helpers;
using SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderCollection.Renderers
{
    public class ClassRenderer : IRenderer
    {
        public ClassRenderer(ITemplateCollection templateCollection, IRendererCollection rendererCollection)
        {
            this.templateCollection = templateCollection;
            this.rendererCollection = rendererCollection;
        }

        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            var lastRow = template.Content.Cells.LastOrDefault();
            foreach(var row in template.Content.Cells)
            {
                foreach(var cell in row)
                {
                    tableBuilder.PushState(new Styler(cell));

                    var childModel = ExtractChildModel(model, cell);

                    if(TemplateDescriptionHelper.Instance.IsCorrectFormValueDescription(cell.StringValue))
                    {
                        childModel = StrictExtractChildModel(model, cell);
                        if(childModel != null)
                        {
                            var typeName = TemplateDescriptionHelper.Instance.GetFormControlTypeFromValueDescription(cell.StringValue);
                            var name = TemplateDescriptionHelper.Instance.GetFormControlNameFromValueDescription(cell.StringValue);
                            var renderer = rendererCollection.GetFormControlRenderer(typeName, childModel.GetType());
                            renderer.Render(tableBuilder, name, childModel);
                        }
                        tableBuilder.SetCurrentStyle();
                        tableBuilder.MoveToNextColumn();
                    }
                    else
                    {
                        var childTemplateName = ExtractTemplateName(cell);
                        var renderer = rendererCollection.GetRenderer(childModel.GetType());
                        renderer.Render(tableBuilder, childModel, templateCollection.GetTemplate(childTemplateName));
                    }

                    tableBuilder.PopState();
                }
                if(!row.Equals(lastRow))
                    tableBuilder.MoveToNextLayer();
            }
            MergeCells(tableBuilder, template);
            ResizeColumns(tableBuilder, template);
        }

        private static void ResizeColumns(ITableBuilder tableBuilder, RenderingTemplate template)
        {
            foreach(var column in template.Columns)
            {
                tableBuilder.ExpandColumn(column.Index - template.Range.UpperLeft.ColumnIndex + 1,
                                          column.Width);
            }
        }

        private static void MergeCells(ITableBuilder tableBuilder, RenderingTemplate template)
        {
            foreach(var mergedCells in template.MergedCells)
                tableBuilder.MergeCells(mergedCells);
        }

        private static string ExtractTemplateName(ICell cell)
        {
            return TemplateDescriptionHelper.Instance.GetTemplateNameFromValueDescription(cell.StringValue);
        }

        private static object ExtractChildModel(object model, ICell cell)
        {
            var expression = cell.StringValue;
            return ExtractChildIfCorrectDescription(expression, model, "") ?? expression ?? "";
        }

        private static object StrictExtractChildModel(object model, ICell cell)
        {
            var expression = cell.StringValue;
            return ExtractChildIfCorrectDescription(expression, model, null);
        }

        private static object ExtractChildIfCorrectDescription(string expression, object model, object onNullReplacement)
        {
            if(TemplateDescriptionHelper.Instance.IsCorrectFormValueDescription(expression) || TemplateDescriptionHelper.Instance.IsCorrectValueDescription(expression))
            {
                var excelTemplateExpression = new ExcelTemplateExpression(expression);
                try
                {
                    return ObjectPropertiesExtractor.ExtractChildObject(model, excelTemplateExpression) ?? onNullReplacement;
                }
                catch(ObjectPropertyExtractionException exception)
                {
                    throw new InvalidExcelTemplateException($"Failed to extract child by path '{excelTemplateExpression.ChildObjectPath.RawPath}' in model of type {model.GetType()}", exception);
                }
            }
            return null;
        }

        private readonly ITemplateCollection templateCollection;
        private readonly IRendererCollection rendererCollection;
    }
}