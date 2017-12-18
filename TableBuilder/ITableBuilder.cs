using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public interface ITableBuilder
    {
        ITableBuilder RenderAtomicValue(string value);
        ITableBuilder RenderAtomicValue(int value);
        ITableBuilder RenderAtomicValue(double value);
        ITableBuilder RenderAtomicValue(decimal value);
        ITableBuilder RenderCheckBoxValue(string name, bool value);
        ITableBuilder RenderDropDownValue(string name, string value);
        ITableBuilder PushState(ICellPosition newOrigin, IStyler styler);
        ITableBuilder PushState(IStyler styler);
        ITableBuilder PushState();
        ITableBuilder PopState();
        ITableBuilder MoveToNextLayer();
        ITableBuilder MoveToNextColumn();
        ITableBuilder ExpandColumn(int relativeColumnIndex, double width);
        ITableBuilder SetCurrentStyle();
        ITableBuilder MergeCells(IRectangle rectangle);
        ITableBuilder AddFormControlInfos(IExcelFormControlInfo[] excelFormControlInfo);
        TableNavigatorState CurrentState { get; }
    }
}