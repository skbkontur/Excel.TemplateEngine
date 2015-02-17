using System;
using System.Collections.Generic;
using System.Globalization;

using SKBKontur.Catalogue.ExcelObjectPrinter.DataTypes;
using SKBKontur.Catalogue.ExcelObjectPrinter.DocumentPrimitivesInterfaces;
using SKBKontur.Catalogue.ExcelObjectPrinter.NavigationPrimitives;

using log4net;

namespace SKBKontur.Catalogue.ExcelObjectPrinter.TableBuilder
{
    public class TableBuilder : ITableBuilder
    {
        public TableBuilder(ITable target, ICellPosition startPosition)
        {
            this.target = target;
            var initialState = new TableBuilderState
                {
                    Origin = startPosition,
                    Cursor = startPosition,
                    CurrentLayerStartRowIndex = startPosition.RowIndex
                };
            states = new Stack<TableBuilderState>(new[] {initialState});
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

        public ITableBuilder PushState(ICellPosition newOrigin)
        {
            var newState = new TableBuilderState
                {
                    Origin = newOrigin,
                    Cursor = newOrigin,
                    CurrentLayerStartRowIndex = newOrigin.RowIndex
                };
            states.Push(newState);
            return this;
        }

        public ITableBuilder PushState()
        {
            return PushState(CurrentState.Cursor);
        }

        public ITableBuilder PopState()
        {
            if(states.Count == 1)
            {
                logger.Warn("Unexpected attempt to pop state.");
                return this;
            }

            var childState = states.Peek();
            states.Pop();
            CurrentState.Cursor = CurrentState.Cursor.Add(new ObjectSize(childState.GlobalWidth, 0));
            CurrentState.CurrentLayerHeight = Math.Max(CurrentState.CurrentLayerHeight, childState.GlobalHeight);
            UpdateCurrentState();
            return this;
        }

        public ITableBuilder MoveToNextLayer()
        {
            var nextLayerStartRowIndex = CurrentState.CurrentLayerStartRowIndex + CurrentState.CurrentLayerHeight + 1;
            CurrentState.Cursor = new CellPosition(nextLayerStartRowIndex, CurrentState.Origin.ColumnIndex);
            CurrentState.CurrentLayerStartRowIndex = nextLayerStartRowIndex;
            CurrentState.CurrentLayerHeight = 0;
            UpdateCurrentState();
            return this;
        }

        public TableBuilderState CurrentState { get { return states.Peek(); } }

        private ITableBuilder RenderAtomicValue(string value, CellType cellType)
        {
            var cell = target.InsertCell(CurrentState.Cursor);
            cell.StringValue = value;
            cell.CellType = cellType;
            CurrentState.Cursor = CurrentState.Cursor.Add(new ObjectSize(1, 0));
            UpdateCurrentState();
            return this;
        }

        private void UpdateCurrentState()
        {
            var currentHeightInLayer = CurrentState.Cursor.RowIndex - CurrentState.CurrentLayerStartRowIndex;
            CurrentState.CurrentLayerHeight = Math.Max(CurrentState.CurrentLayerHeight, currentHeightInLayer);
            var currentCursorOffset = CurrentState.Cursor.Subtract(CurrentState.Origin);
            CurrentState.GlobalHeight = Math.Max(CurrentState.CurrentLayerHeight,
                                                 Math.Max(CurrentState.GlobalHeight, currentCursorOffset.Height));
            CurrentState.GlobalWidth = Math.Max(CurrentState.GlobalWidth, currentCursorOffset.Width);
        }

        private readonly Stack<TableBuilderState> states;
        private readonly ITable target;
        private readonly ILog logger = LogManager.GetLogger(typeof(TableBuilder));
    }
}