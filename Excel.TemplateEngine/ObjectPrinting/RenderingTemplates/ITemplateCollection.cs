namespace Excel.TemplateEngine.ObjectPrinting.RenderingTemplates
{
    public interface ITemplateCollection
    {
        RenderingTemplate GetTemplate(string templateName);
    }
}