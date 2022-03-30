using System.Globalization;

using JetBrains.Annotations;

using SkbKontur.Excel.TemplateEngine.ObjectPrinting.ExcelDocumentPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.NavigationPrimitives;
using SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableNavigator;

namespace SkbKontur.Excel.TemplateEngine.ObjectPrinting.TableParser
{
    public class TableParser : ITableParser
    {
        public TableParser([NotNull] ITable target, [NotNull] ITableNavigator navigator)
        {
            this.target = target;
            this.navigator = navigator;
        }

        public bool TryParseCheckBoxValue([NotNull] string name, out bool result)
        {
            var formControl = target.TryGetCheckBoxFormControl(name);
            if (formControl == null)
            {
                result = false;
                return false;
            }
            result = formControl.IsChecked;
            return true;
        }

        public bool TryParseDropDownValue([NotNull] string name, [CanBeNull] out string result)
        {
            var formControl = target.TryGetDropDownFormControl(name);
            if (formControl == null)
            {
                result = null;
                return false;
            }
            result = formControl.SelectedValue;
            return true;
        }

        [NotNull]
        public ITableParser PushState(ICellPosition newOrigin)
        {
            navigator.PushState(newOrigin);
            return this;
        }

        [NotNull]
        public ITableParser PushState()
        {
            navigator.PushState();
            return this;
        }

        [NotNull]
        public ITableParser PopState()
        {
            navigator.PopState();
            return this;
        }

        [NotNull]
        public ITableParser MoveToNextLayer()
        {
            navigator.MoveToNextLayer();
            return this;
        }

        [NotNull]
        public ITableParser MoveToNextColumn()
        {
            navigator.MoveToNextColumn();
            return this;
        }

        [CanBeNull]
        public string GetCurrentCellText()
        {
            return target.GetCell(CurrentState.Cursor)?.StringValue;
        }

        public TableNavigatorState CurrentState => navigator.CurrentState;

        private const NumberStyles numberStyles = NumberStyles.AllowDecimalPoint | NumberStyles.AllowLeadingSign;

        private readonly ITable target;
        private readonly ITableNavigator navigator;
        private static readonly CultureInfo russianCultureInfo = CultureInfo.GetCultureInfo("ru-RU");
    }
}