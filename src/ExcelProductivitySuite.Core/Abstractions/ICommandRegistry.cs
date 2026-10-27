using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Core.Abstractions
{
    public interface ICommandRegistry
    {
        void RegisterCommand(SuiteCommand command);
        SuiteCommand? GetCommand(string id);
        IReadOnlyList<SuiteCommand> GetAllCommands();
        IReadOnlyList<SuiteCommand> SearchCommands(string query, CommandCategory? categoryFilter = null);
        void Execute(string commandId, IExcelContext context);
    }

    /// <summary>
    /// Scoped state manager that suppresses ScreenUpdating, Events, and Calculation
    /// during high-performance bulk operations and strictly restores them on Dispose.
    /// </summary>
    public interface IExcelStateScope : IDisposable
    {
        void SuppressScreenUpdating();
        void SuppressEvents();
        void SetManualCalculation();
        void SuppressAlerts();
    }

    public interface ILogger
    {
        void LogInfo(string message);
        void LogWarning(string message);
        void LogError(string message, Exception? ex = null);
    }
}
