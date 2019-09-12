namespace Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations
{
    public class ObjectSize : IObjectSize
    {
        public ObjectSize(int width, int height)
        {
            Width = width;
            Height = height;
        }

        public int Width { get; }
        public int Height { get; }

        public IObjectSize Add(IObjectSize other)
        {
            return new ObjectSize(Width + other.Width, Height + other.Height);
        }

        public IObjectSize Subtract(IObjectSize other)
        {
            return new ObjectSize(Width - other.Width, Height - other.Height);
        }
    }
}