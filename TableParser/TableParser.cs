using System;
using System.Globalization;

using SKBKontur.Catalogue.ExcelFileGenerator.Interfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;
using SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableParser
{
    public class TableParser : ITableParser
    {
        private readonly ITableNavigator navigator;

        public TableParser(ITableNavigator navigator)
        {
            this.navigator = navigator;
        }

        /// <summary>
        /// Be careful! Result can be both null and "" when cell is empty (seems like it depends on the way of file creation)
        /// </summary>
        /// <param name="result"></param>
        /// <returns></returns>
        public bool TryParseAtomicValue(out string result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            result = cellValue;
            return result != null;
        }

        public bool TryParseAtomicValue(out int result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            return int.TryParse(cellValue, out result);
        }

        public bool TryParseAtomicValue(out double result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            return double.TryParse(cellValue, out result);
        }

        public bool TryParseAtomicValue(out decimal result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            return decimal.TryParse(cellValue, numberStyles, russianCultureInfo, out result) || decimal.TryParse(cellValue, numberStyles, CultureInfo.InvariantCulture, out result);
        }

        public bool TryParseAtomicValue(out long result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            return long.TryParse(cellValue, out result);
        }

        public bool TryParseAtomicValue(out int? result)
        {
            return TryParseNullableAtomicValue(() => (TryParseAtomicValue(out int res), res), out result);
        }

        public bool TryParseAtomicValue(out double? result)
        {
            return TryParseNullableAtomicValue(() => (TryParseAtomicValue(out double res), res), out result);
        }

        public bool TryParseAtomicValue(out decimal? result)
        {
            return TryParseNullableAtomicValue(() => (TryParseAtomicValue(out decimal res), res), out result);
        }

        public bool TryParseAtomicValue(out long? result)
        {
            return TryParseNullableAtomicValue(() => (TryParseAtomicValue(out long res), res), out result);
        }

        public bool TryParseNullableAtomicValue<T>(Func<(bool succeed, T result)> parser, out T? result)
            where T : struct
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            if (string.IsNullOrEmpty(cellValue))
            {
                result = null;
                return true;
            }
            bool succeed;
            (succeed, result) = parser();
            return succeed;
        }

        public bool TryParseCheckBoxValue(string name, out bool result)
        {
            var formControl = Target.TryGetFormControl<IExcelCheckBoxControlInfo>(name);
            if(formControl == null)
            {
                result = false;
                return false;
            }
            result = formControl.IsChecked;
            return true;
        }

        public bool TryParseDropDownValue(string name, out string result)
        {
            var formControl = Target.TryGetFormControl<IExcelDropDownControlInfo>(name);
            if (formControl == null)
            {
                result = null;
                return false;
            }
            result = formControl.SelectedValue;
            return true;
        }

        public ITableParser PushState(ICellPosition newOrigin, IStyler styler)
        {
            navigator.PushState(newOrigin, styler);
            return this;
        }

        public ITableParser PushState(IStyler styler)
        {
            navigator.PushState(styler);
            return this;
        }

        public ITableParser PushState()
        {
            navigator.PushState();
            return this;
        }

        public ITableParser PopState()
        {
            navigator.PopState();
            return this;
        }

        public ITableParser MoveToNextLayer()
        {
            navigator.MoveToNextLayer();
            return this;
        }

        public ITableParser MoveToNextColumn()
        {
            navigator.MoveToNextColumn();
            return this;
        }

        public ITableParser SetCurrentStyle()
        {
            navigator.SetCurrentStyle();
            return this;
        }

        public TableNavigatorState CurrentState => navigator.CurrentState;
        private ITable Target => navigator.Target;

        private const NumberStyles numberStyles = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign;
        private static readonly CultureInfo russianCultureInfo = CultureInfo.GetCultureInfo("ru-RU");
    }
}