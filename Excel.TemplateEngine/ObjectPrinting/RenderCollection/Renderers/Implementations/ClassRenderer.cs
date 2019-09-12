using System.Linq;

using SkbKontur.Excel.TemplateEngine.Exceptions;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.Helpers;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderingTemplates;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableBuilder;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.RenderCollection.Renderers.Implementations
{
    internal class ClassRenderer : IRenderer
    {
        public ClassRenderer(ITemplateCollection templateCollection, IRendererCollection rendererCollection)
        {
            this.templateCollection = templateCollection;
            this.rendererCollection = rendererCollection;
        }

        public void Render(ITableBuilder tableBuilder, object model, RenderingTemplate template)
        {
            var lastRow = template.Content.Cells.LastOrDefault();
            foreach (var row in template.Content.Cells)
            {
                foreach (var cell in row)
                {
                    tableBuilder.PushState(new Style(cell));

                    if (TemplateDescriptionHelper.IsCorrectFormValueDescription(cell.StringValue))
                        RenderFormControl(tableBuilder, model, cell);
                    else
                        RenderCellularValue(tableBuilder, model, cell);

                    tableBuilder.PopState();
                }
                if (!row.Equals(lastRow))
                    tableBuilder.MoveToNextLayer();
            }
            MergeCells(tableBuilder, template);
            ResizeColumns(tableBuilder, template);
        }

        private void RenderCellularValue(ITableBuilder tableBuilder, object model, ICell cell)
        {
            var childModel = ExtractChildModel(model, cell);
            var childTemplateName = ExtractTemplateName(cell);
            var renderer = rendererCollection.GetRenderer(childModel.GetType());
            renderer.Render(tableBuilder, childModel, templateCollection.GetTemplate(childTemplateName));
        }

        private void RenderFormControl(ITableBuilder tableBuilder, object model, ICell cell)
        {
            var childModel = StrictExtractChildModel(model, cell);
            if (childModel != null)
            {
                var (controlType, controlName) = TemplateDescriptionHelper.TryGetFormControlFromValueDescription(cell.StringValue);
                var renderer = rendererCollection.GetFormControlRenderer(controlType, childModel.GetType());
                renderer.Render(tableBuilder, controlName, childModel);
            }
            tableBuilder.SetCurrentStyle();
            tableBuilder.MoveToNextColumn();
        }

        private static void ResizeColumns(ITableBuilder tableBuilder, RenderingTemplate template)
        {
            foreach (var column in template.Columns)
            {
                tableBuilder.ExpandColumn(column.Index - template.Range.UpperLeft.ColumnIndex + 1,
                                          column.Width);
            }
        }

        private static void MergeCells(ITableBuilder tableBuilder, RenderingTemplate template)
        {
            foreach (var mergedCells in template.MergedCells)
                tableBuilder.MergeCells(mergedCells);
        }

        private static string ExtractTemplateName(ICell cell)
        {
            return TemplateDescriptionHelper.GetTemplateNameFromValueDescription(cell.StringValue);
        }

        private static object ExtractChildModel(object model, ICell cell)
        {
            var expression = cell.StringValue;
            if (!TemplateDescriptionHelper.IsCorrectValueDescription(expression))
                return expression ?? "";
            return ExtractChildIfCorrectDescription(expression, model) ?? "";
        }

        private static object StrictExtractChildModel(object model, ICell cell)
        {
            var expression = cell.StringValue;
            return ExtractChildIfCorrectDescription(expression, model);
        }

        private static object ExtractChildIfCorrectDescription(string expression, object model)
        {
            var excelTemplatePath = ExcelTemplatePath.FromRawExpression(expression);
            try
            {
                return ObjectPropertiesExtractor.ExtractChildObject(model, excelTemplatePath);
            }
            catch (ObjectPropertyExtractionException exception)
            {
                throw new ExcelTemplateEngineException($"Failed to extract child by path '{excelTemplatePath.RawPath}' in model of type {model.GetType()}", exception);
            }
        }

        private readonly ITemplateCollection templateCollection;
        private readonly IRendererCollection rendererCollection;
    }
}