using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Engine.Recovery
{
    /// <summary>
    /// Thread-safe multi-level Undo/Redo manager tracking cell and range state transitions
    /// with no-op filtering, workbook isolation, and capacity bounds.
    /// </summary>
    public class UndoManager : IUndoManager
    {
        private readonly object _lock = new();
        private readonly List<UndoStep> _undoStack;
        private readonly Stack<UndoStep> _redoStack;
        private readonly int _maxHistory;

        public UndoManager(int maxHistory = 50)
        {
            _maxHistory = Math.Max(5, maxHistory);
            _undoStack = new List<UndoStep>(_maxHistory);
            _redoStack = new Stack<UndoStep>(_maxHistory);
        }

        public bool CanUndo
        {
            get
            {
                lock (_lock)
                {
                    return _undoStack.Count > 0;
                }
            }
        }

        public bool CanRedo
        {
            get
            {
                lock (_lock)
                {
                    return _redoStack.Count > 0;
                }
            }
        }

        public int UndoCount
        {
            get
            {
                lock (_lock)
                {
                    return _undoStack.Count;
                }
            }
        }

        public int RedoCount
        {
            get
            {
                lock (_lock)
                {
                    return _redoStack.Count;
                }
            }
        }

        public string? LastUndoDescription
        {
            get
            {
                lock (_lock)
                {
                    return _undoStack.Count > 0 ? _undoStack[_undoStack.Count - 1].Description : null;
                }
            }
        }

        public string? LastRedoDescription
        {
            get
            {
                lock (_lock)
                {
                    return _redoStack.Count > 0 ? _redoStack.Peek().Description : null;
                }
            }
        }

        public void RecordAction(UndoStep step)
        {
            if (step == null) throw new ArgumentNullException(nameof(step));

            // Discard no-op actions where data did not actually change
            if (AreMatricesEqual(step.BeforeValues, step.AfterValues) &&
                AreMatricesEqual(step.BeforeFormulas, step.AfterFormulas))
            {
                return;
            }

            lock (_lock)
            {
                if (_undoStack.Count >= _maxHistory)
                {
                    _undoStack.RemoveAt(0); // Discard oldest
                }

                _undoStack.Add(step);
                _redoStack.Clear(); // Clearing redo stack on new operation
            }
        }

        public void RecordAction(
            string description,
            string sheetName,
            string rangeAddress,
            int startRow,
            int startColumn,
            object?[,] beforeValues,
            object?[,] afterValues,
            object?[,]? beforeFormulas = null,
            object?[,]? afterFormulas = null)
        {
            RecordAction(
                description,
                string.Empty,
                null,
                sheetName,
                rangeAddress,
                startRow,
                startColumn,
                beforeValues,
                afterValues,
                beforeFormulas,
                afterFormulas);
        }

        public void RecordAction(
            string description,
            string workbookName,
            string? workbookPath,
            string sheetName,
            string rangeAddress,
            int startRow,
            int startColumn,
            object?[,] beforeValues,
            object?[,] afterValues,
            object?[,]? beforeFormulas = null,
            object?[,]? afterFormulas = null)
        {
            var step = new UndoStep(
                description,
                workbookName,
                workbookPath,
                sheetName,
                rangeAddress,
                startRow,
                startColumn,
                beforeValues,
                afterValues,
                beforeFormulas,
                afterFormulas);

            RecordAction(step);
        }

        public UndoStep? PopUndo()
        {
            lock (_lock)
            {
                if (_undoStack.Count == 0) return null;

                int lastIdx = _undoStack.Count - 1;
                var step = _undoStack[lastIdx];
                _undoStack.RemoveAt(lastIdx);
                return step;
            }
        }

        public UndoStep? PopRedo()
        {
            lock (_lock)
            {
                if (_redoStack.Count == 0) return null;
                return _redoStack.Pop();
            }
        }

        public UndoStep? PeekUndo()
        {
            lock (_lock)
            {
                if (_undoStack.Count == 0) return null;
                return _undoStack[_undoStack.Count - 1];
            }
        }

        public UndoStep? PeekRedo()
        {
            lock (_lock)
            {
                if (_redoStack.Count == 0) return null;
                return _redoStack.Peek();
            }
        }

        public void PushRedo(UndoStep step)
        {
            if (step == null) throw new ArgumentNullException(nameof(step));

            lock (_lock)
            {
                _redoStack.Push(step);
            }
        }

        public void PushUndo(UndoStep step)
        {
            if (step == null) throw new ArgumentNullException(nameof(step));

            lock (_lock)
            {
                if (_undoStack.Count >= _maxHistory)
                {
                    _undoStack.RemoveAt(0);
                }

                _undoStack.Add(step);
            }
        }

        public void Clear()
        {
            lock (_lock)
            {
                _undoStack.Clear();
                _redoStack.Clear();
            }
        }

        public void ClearForWorkbook(string workbookName)
        {
            if (string.IsNullOrWhiteSpace(workbookName)) return;

            lock (_lock)
            {
                _undoStack.RemoveAll(s => string.Equals(s.WorkbookName, workbookName, StringComparison.OrdinalIgnoreCase));

                var remainingRedo = new List<UndoStep>(_redoStack);
                remainingRedo.RemoveAll(s => string.Equals(s.WorkbookName, workbookName, StringComparison.OrdinalIgnoreCase));
                remainingRedo.Reverse(); // Preserve LIFO order when rebuilding stack

                _redoStack.Clear();
                foreach (var step in remainingRedo)
                {
                    _redoStack.Push(step);
                }
            }
        }

        public static bool AreMatricesEqual(object?[,]? a, object?[,]? b)
        {
            if (ReferenceEquals(a, b)) return true;
            if (a == null || b == null) return false;

            int r1 = a.GetLength(0), c1 = a.GetLength(1);
            int r2 = b.GetLength(0), c2 = b.GetLength(1);
            if (r1 != r2 || c1 != c2) return false;

            int aRStart = a.GetLowerBound(0), aCStart = a.GetLowerBound(1);
            int bRStart = b.GetLowerBound(0), bCStart = b.GetLowerBound(1);

            for (int r = 0; r < r1; r++)
            {
                for (int c = 0; c < c1; c++)
                {
                    object? valA = a[aRStart + r, aCStart + c];
                    object? valB = b[bRStart + r, bCStart + c];

                    if (valA == null && valB == null) continue;
                    if (valA == null || valB == null) return false;

                    // Numeric tolerance comparison for floating-point calculations
                    if (valA is double da && valB is double db)
                    {
                        if (Math.Abs(da - db) > 1e-9) return false;
                        continue;
                    }

                    if (!Equals(valA, valB)) return false;
                }
            }

            return true;
        }
    }
}
