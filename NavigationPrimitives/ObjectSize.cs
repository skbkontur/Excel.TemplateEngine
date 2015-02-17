namespace SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives
{
    public class ObjectSize : IObjectSize
    {
        public ObjectSize(int width, int height)
        {
            Width = width;
            Height = height;
        }

        public int Width { get; private set; }
        public int Height { get; private set; }

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