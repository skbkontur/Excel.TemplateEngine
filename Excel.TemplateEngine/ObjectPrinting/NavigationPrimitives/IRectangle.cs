using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives.Implementations;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives
{
    public interface IRectangle
    {
        ICellPosition UpperLeft { get; }
        ICellPosition LowerRight { get; }
        IObjectSize Size { get; }
        bool Intersects(IRectangle rect);
        bool Contains(ICellPosition position);
        bool Contains(IRectangle rect);
    }

    public static class RectangleExtensions
    {
        public static IRectangle ToRelativeCoordinates(this IRectangle globalCoordinates, ICellPosition newOrigin)
        {
            return new Rectangle(globalCoordinates.UpperLeft.ToRelativeCoordinates(newOrigin),
                                 globalCoordinates.LowerRight.ToRelativeCoordinates(newOrigin));
        }

        public static IRectangle ToGlobalCoordinates(this IRectangle relativeCoordinates, ICellPosition originCoordinates)
        {
            return new Rectangle(relativeCoordinates.UpperLeft.ToGlobalCoordinates(originCoordinates),
                                 relativeCoordinates.LowerRight.ToGlobalCoordinates(originCoordinates));
        }
    }
}