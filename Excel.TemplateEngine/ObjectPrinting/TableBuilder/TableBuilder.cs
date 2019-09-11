using System.Collections.Generic;
using System.Globalization;
using System.Linq;

using Excel.TemplateEngine.Exceptions;
using Excel.TemplateEngine.ObjectPrinting.DataTypes;
using Excel.TemplateEngine.ObjectPrinting.DocumentPrimitivesInterfaces;
using Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using Excel.TemplateEngine.ObjectPrinting.TableNavigator;

using JetBrains.Annotations;

namespace Excel.TemplateEngine.ObjectPrinting.TableBuilder
{
    public class TableBuilder : ITableBuilder
    {
        public TableBuilder([NotNull] ITable target, [NotNull] ITableNavigator navigator, [CanBeNull] IStyle style = null)
        {
            this.target = target;
            this.navigator = navigator;
            styles = new Stack<IStyle>(new[] {style});
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

        [NotNull]
        public ITableBuilder RenderCheckBoxValue([NotNull] string name, bool value)
        {
            var formControl = target.TryGetCheckBoxFormControl(name);
            if (formControl == null)
                throw new ExcelTemplateEngineException($"CheckBox with name {name} not found");
            formControl.IsChecked = value;
            return this;
        }

        [NotNull]
        public ITableBuilder RenderDropDownValue([NotNull] string name, [CanBeNull] string value)
        {
            var formControl = target.TryGetDropDownFormControl(name);
            if (formControl == null)
                throw new ExcelTemplateEngineException($"DropDown with name {name} not found");
            formControl.SelectedValue = value;
            return this;
        }

        public ITableBuilder PushState(ICellPosition newOrigin, IStyle style)
        {
            navigator.PushState(newOrigin);
            styles.Push(style);
            return this;
        }

        public ITableBuilder PushState(IStyle style)
        {
            navigator.PushState();
            styles.Push(style);
            return this;
        }

        public ITableBuilder PushState()
        {
            navigator.PushState();
            styles.Push(CurrentStyle);
            return this;
        }

        public ITableBuilder PopState()
        {
            navigator.PopState();
            styles.Pop();
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
            var cell = target.GetCell(navigator.CurrentState.Cursor) ?? target.InsertCell(navigator.CurrentState.Cursor);
            CurrentStyle.ApplyTo(cell);
            return this;
        }

        public ITableBuilder ExpandColumn(int relativeColumnIndex, double width)
        {
            var globalIndex = relativeColumnIndex + navigator.CurrentState.Origin.ColumnIndex - 1;
            var currentWidth = target.Columns
                                     .FirstOrDefault(col => col.Index == globalIndex)?.Width ?? 0.0;

            if (currentWidth < width)
                target.ResizeColumn(globalIndex, width);
            return this;
        }

        public ITableBuilder MergeCells(IRectangle rectangle)
        {
            target.MergeCells(rectangle.ToGlobalCoordinates(navigator.CurrentState.Origin));
            return this;
        }

        public ITableBuilder CopyFormControlsFrom([NotNull] ITable template)
        {
            target.CopyFormControlsFrom(template);
            return this;
        }

        public ITableBuilder CopyDataValidationsFrom([NotNull] ITable template)
        {
            target.CopyDataValidationsFrom(template);
            return this;
        }

        public ITableBuilder CopyWorksheetExtensionListFrom([NotNull] ITable template)
        {
            target.CopyWorksheetExtensionListFrom(template);
            return this;
        }

        public ITableBuilder CopyCommentsFrom([NotNull] ITable template)
        {
            target.CopyCommentsFrom(template);
            return this;
        }

        private ITableBuilder RenderAtomicValue(string value, CellType cellType)
        {
            var cell = target.GetCell(navigator.CurrentState.Cursor) ?? target.InsertCell(navigator.CurrentState.Cursor);
            cell.StringValue = value;
            cell.CellType = cellType;
            return this;
        }

        public TableNavigatorState CurrentState => navigator.CurrentState;
        private IStyle CurrentStyle => styles.Peek();
        private readonly ITable target;
        private readonly ITableNavigator navigator;
        private readonly Stack<IStyle> styles;
    }
}