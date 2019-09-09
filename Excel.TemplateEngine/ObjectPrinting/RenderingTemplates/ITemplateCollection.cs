namespace SKBKontur.Catalogue.ExcelObjectPrinter.RenderingTemplates
{
    public interface ITemplateCollection
    {
        RenderingTemplate GetTemplate(string templateName);
    }
}