using System;
using System.Collections.Generic;

using log4net;

using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableNavigator
{
    public class TableNavigator : ITableNavigator
    {
        public TableNavigator(ITable target, ICellPosition startPosition, IStyler styler = null)
        {
            Target = target;
            var initialState = new TableNavigatorState
                {
                    Origin = startPosition,
                    Cursor = startPosition,
                    CurrentLayerStartRowIndex = startPosition.RowIndex,
                    Styler = styler
                };
            states = new Stack<TableNavigatorState>(new[] {initialState});
        }

        public void PushState(ICellPosition newOrigin, IStyler styler)
        {
            var newState = new TableNavigatorState
                {
                    Origin = newOrigin,
                    Cursor = newOrigin,
                    CurrentLayerStartRowIndex = newOrigin.RowIndex,
                    Styler = styler
                };
            states.Push(newState);
        }

        public void PushState(IStyler styler)
        {
            PushState(CurrentState.Cursor, styler);
        }

        public void PushState()
        {
            PushState(CurrentState.Cursor, CurrentState.Styler);
        }

        public void PopState()
        {
            if(states.Count == 1)
            {
                logger.Warn("Unexpected attempt to pop state.");
                return;
            }

            var childState = states.Peek();
            states.Pop();
            CurrentState.Cursor = CurrentState.Cursor.Add(new ObjectSize(childState.GlobalWidth, 0));
            CurrentState.CurrentLayerHeight = Math.Max(CurrentState.CurrentLayerHeight, childState.GlobalHeight);
            UpdateCurrentState();
        }

        public void MoveToNextLayer()
        {
            var nextLayerStartRowIndex = CurrentState.CurrentLayerStartRowIndex + CurrentState.CurrentLayerHeight + 1;
            CurrentState.Cursor = new CellPosition(nextLayerStartRowIndex, CurrentState.Origin.ColumnIndex);
            CurrentState.CurrentLayerStartRowIndex = nextLayerStartRowIndex;
            CurrentState.CurrentLayerHeight = 0;
            UpdateCurrentState();
        }

        public void MoveToNextColumn()
        {
            CurrentState.Cursor = CurrentState.Cursor.Add(new ObjectSize(1, 0));
            UpdateCurrentState();
        }

        public void SetCurrentStyle()
        {
            var cell = Target.GetCell(CurrentState.Cursor) ?? Target.InsertCell(CurrentState.Cursor);
            CurrentState.Styler.ApplyStyle(cell);
        }

        public TableNavigatorState CurrentState => states.Peek();
        public ITable Target { get; }

        private void UpdateCurrentState()
        {
            var currentHeightInLayer = CurrentState.Cursor.RowIndex - CurrentState.CurrentLayerStartRowIndex;
            CurrentState.CurrentLayerHeight = Math.Max(CurrentState.CurrentLayerHeight, currentHeightInLayer);
            var currentCursorOffset = CurrentState.Cursor.Subtract(CurrentState.Origin);
            CurrentState.GlobalHeight = Math.Max(CurrentState.CurrentLayerHeight,
                                                 Math.Max(CurrentState.GlobalHeight, currentCursorOffset.Height));
            CurrentState.GlobalWidth = Math.Max(CurrentState.GlobalWidth, currentCursorOffset.Width);
        }

        private readonly Stack<TableNavigatorState> states;
        private readonly ILog logger = LogManager.GetLogger(typeof(TableNavigator));
    }
}