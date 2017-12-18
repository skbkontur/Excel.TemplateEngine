using System;
using System.Globalization;
using System.Linq;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.DataTypes;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.ExcelDocumentPrimitivesImplementation;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;
using SKBKontur.Catalogue.Objects;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public class TableBuilder : ITableBuilder
    {
        private readonly ITableNavigator navigator;
        
        public TableBuilder(ITableNavigator navigator)
        {
            this.navigator = navigator;
        }

        public ITableBuilder RenderAtomicValue(string value)
        {
            return RenderAtomicValue(value, CellType.String);
        }

        public ITableBuilder RenderAtomicValue(int value)
        {
            return RenderAtomicValue(value.ToString(CultureInfo.InvariantCulture), CellType.Number);
        }

        public ITableBuilder RenderAtomicValue(double value)
        {
            return RenderAtomicValue(value.ToString(CultureInfo.InvariantCulture), CellType.Number);
        }

        public ITableBuilder RenderAtomicValue(decimal value)
        {
            return RenderAtomicValue(value.ToString(CultureInfo.InvariantCulture), CellType.Number);
        }

        public ITableBuilder RenderCheckBoxValue(string name, bool value)
        {
            var formControl = Target.TryGetFormControl(name);
            if (formControl == null)
                throw new ArgumentException($"Form control with name {name} not found");
            formControl.ExcelFormControlInfo.IsChecked = value;
            return this;
        }

        public ITableBuilder RenderDropDownValue(string name, string value)
        {
            var formControl = Target.TryGetFormControl(name);
            if (formControl == null)
                throw new ArgumentException($"Form control with name {name} not found");
            formControl.ExcelFormControlInfo.SelectedValue = value;
            return this;
        }

        public ITableBuilder PushState(ICellPosition newOrigin, IStyler styler)
        {
            navigator.PushState(newOrigin, styler);
            return this;
        }

        public ITableBuilder PushState(IStyler styler)
        {
            navigator.PushState(styler);
            return this;
        }

        public ITableBuilder PushState()
        {
            navigator.PushState();
            return this;
        }

        public ITableBuilder PopState()
        {
            navigator.PopState();
            return this;
        }

        public ITableBuilder MoveToNextLayer()
        {
            navigator.MoveToNextLayer();
            return this;
        }

        public ITableBuilder MoveToNextColumn()
        {
            navigator.MoveToNextColumn();
            return this;
        }

        public ITableBuilder SetCurrentStyle()
        {
            navigator.SetCurrentStyle();
            return this;
        }

        public ITableBuilder ExpandColumn(int relativeColumnIndex, double width)
        {
            var globalIndex = relativeColumnIndex + CurrentState.Origin.ColumnIndex - 1;
            var currentWidth = Target.Columns
                                     .FirstOrDefault(col => col.Index == globalIndex)?.Width ?? 0.0;

            if(currentWidth < width)
                Target.ResizeColumn(globalIndex, width);
            return this;
        }
        

        public ITableBuilder MergeCells(IRectangle rectangle)
        {
            Target.MergeCells(rectangle.ToGlobalCoordinates(CurrentState.Origin));
            return this;
        }

        public ITableBuilder AddFormControlInfos(IExcelFormControlInfo[] excelFormControlInfo)
        {
            Target.AddFormControls(excelFormControlInfo.Select(x => new ExcelFormControl {ExcelFormControlInfo = x}).Cast<IFormControl>().ToArray());
            return this;
        }

        private ITableBuilder RenderAtomicValue(string value, CellType cellType)
        {
            var cell = Target.GetCell(CurrentState.Cursor) ?? Target.InsertCell(CurrentState.Cursor);
            cell.StringValue = value;
            cell.CellType = cellType;
            return this;
        }

        public TableNavigatorState CurrentState => navigator.CurrentState;
        private ITable Target => navigator.Target;
    }
}
