using System.Globalization;

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

            if(decimal.TryParse(cellValue, numberStyles, russianCultureInfo, out result) || decimal.TryParse(cellValue, numberStyles, CultureInfo.InvariantCulture, out result))
            {
                var s = result.ToString(".0000", CultureInfo.InvariantCulture);
                result = decimal.Parse(s, numberStyles, CultureInfo.InvariantCulture); // todo (mpivko, 15.12.2017): 
                return true;
            }
            return false;
        }

        public bool TryParseAtomicValue(out long result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            return long.TryParse(cellValue, out result);
        }

        public bool TryParseAtomicValue(out int? result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            if(string.IsNullOrEmpty(cellValue))
            {
                result = null;
                return true;
            }
            var succeed = int.TryParse(cellValue, out var intResult);
            result = intResult;
            return succeed;
        }

        public bool TryParseAtomicValue(out double? result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            if (string.IsNullOrEmpty(cellValue))
            {
                result = null;
                return true;
            }
            var succeed = double.TryParse(cellValue, out var doubleResult);
            result = doubleResult;
            return succeed;
        }

        public bool TryParseAtomicValue(out decimal? result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            if (string.IsNullOrEmpty(cellValue))
            {
                result = null;
                return true;
            }
            var succeed = TryParseAtomicValue(out decimal decimalResult);
            result = decimalResult;
            return succeed;
        }

        public bool TryParseAtomicValue(out long? result)
        {
            var cellValue = Target.GetCell(CurrentState.Cursor)?.StringValue;
            if (string.IsNullOrEmpty(cellValue))
            {
                result = null;
                return true;
            }
            var succeed = long.TryParse(cellValue, out var longResult);
            result = longResult;
            return succeed;
        }

        public bool TryParseCheckBoxValue(string name, out bool result)
        {
            var formControl = Target.TryGetFormControl(name);
            if(formControl == null)
            {
                result = false;
                return false;
            }
            result = formControl.ExcelFormControlInfo.IsChecked;
            return true;
        }

        public bool TryParseDropDownValue(string name, out string result)
        {
            var formControl = Target.TryGetFormControl(name);
            if (formControl == null)
            {
                result = null;
                return false;
            }
            result = formControl.ExcelFormControlInfo.SelectedValue;
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