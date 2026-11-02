using System;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Core.Abstractions
{
    /// <summary>
    /// Contract for managing multi-level Undo and Redo operations across Excel Productivity Suite commands.
    /// </summary>
    public interface IUndoManager
    {
        bool CanUndo { get; }
        bool CanRedo { get; }
        int UndoCount { get; }
        int RedoCount { get; }
        string? LastUndoDescription { get; }
        string? LastRedoDescription { get; }

        void RecordAction(UndoStep step);

        void RecordAction(
            string description,
            string sheetName,
            string rangeAddress,
            int startRow,
            int startColumn,
            object?[,] beforeValues,
            object?[,] afterValues,
            object?[,]? beforeFormulas = null,
            object?[,]? afterFormulas = null);

        void RecordAction(
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
            object?[,]? afterFormulas = null);

        UndoStep? PopUndo();

        UndoStep? PopRedo();

        UndoStep? PeekUndo();

        UndoStep? PeekRedo();

        void PushRedo(UndoStep step);

        void PushUndo(UndoStep step);

        void Clear();

        void ClearForWorkbook(string workbookName);
    }
}
