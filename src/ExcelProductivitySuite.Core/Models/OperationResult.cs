using System;

namespace ExcelProductivitySuite.Core.Models
{
    /// <summary>
    /// Represents the execution result of an Excel suite utility.
    /// </summary>
    public class OperationResult
    {
        public bool Success { get; }
        public string Message { get; }
        public int AffectedCellsCount { get; }
        public int SkippedCount { get; }
        public long ElapsedMilliseconds { get; }
        public string? Details { get; }

        public OperationResult(bool success, string message, int affectedCellsCount = 0, int skippedCount = 0, long elapsedMilliseconds = 0, string? details = null)
        {
            Success = success;
            Message = message ?? throw new ArgumentNullException(nameof(message));
            AffectedCellsCount = affectedCellsCount;
            SkippedCount = skippedCount;
            ElapsedMilliseconds = elapsedMilliseconds;
            Details = details;
        }

        public static OperationResult Succeeded(string message, int affectedCells = 0, int skipped = 0, long elapsedMs = 0, string? details = null) =>
            new OperationResult(true, message, affectedCells, skipped, elapsedMs, details);

        public static OperationResult Failed(string message, string? details = null) =>
            new OperationResult(false, message, 0, 0, 0, details);

        public static OperationResult Cancelled(string message = "Operation cancelled by user.") =>
            new OperationResult(false, message);
    }
}
