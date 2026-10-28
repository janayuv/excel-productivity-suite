using System;
using System.Collections.Generic;
using System.Linq;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;

namespace ExcelProductivitySuite.Core.Services
{
    public class CommandRegistry : ICommandRegistry
    {
        private readonly Dictionary<string, SuiteCommand> _commands = new(StringComparer.OrdinalIgnoreCase);

        public void RegisterCommand(SuiteCommand command)
        {
            if (command == null) throw new ArgumentNullException(nameof(command));
            _commands[command.Id] = command;
        }

        public SuiteCommand? GetCommand(string id)
        {
            if (string.IsNullOrWhiteSpace(id)) return null;
            return _commands.TryGetValue(id, out var cmd) ? cmd : null;
        }

        public IReadOnlyList<SuiteCommand> GetAllCommands()
        {
            return _commands.Values.OrderBy(c => c.Category).ThenBy(c => c.Title).ToList();
        }

        public IReadOnlyList<SuiteCommand> SearchCommands(string query, CommandCategory? categoryFilter = null)
        {
            return SearchCommands(query, categoryFilter, null);
        }

        public IReadOnlyList<SuiteCommand> SearchCommands(string query, CommandCategory? categoryFilter, FeatureStatus? statusFilter)
        {
            var source = _commands.Values.AsEnumerable();

            if (categoryFilter.HasValue)
            {
                source = source.Where(c => c.Category == categoryFilter.Value);
            }

            if (statusFilter.HasValue)
            {
                source = source.Where(c => c.Status == statusFilter.Value);
            }

            if (string.IsNullOrWhiteSpace(query))
            {
                return source.OrderBy(c => c.Category).ThenBy(c => c.Title).ToList();
            }

            var cleanQuery = query.Trim();
            return source
                .Where(c => Matches(c, cleanQuery))
                .OrderByDescending(c => Score(c, cleanQuery))
                .ThenBy(c => c.Title)
                .ToList();
        }

        public void Execute(string commandId, IExcelContext context)
        {
            var cmd = GetCommand(commandId);
            if (cmd == null)
            {
                throw new InvalidOperationException($"Command with ID '{commandId}' is not registered in the suite.");
            }

            cmd.Execute(context);
        }

        private static bool Matches(SuiteCommand cmd, string query)
        {
            if (cmd.Title.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (cmd.Description.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (cmd.Category.ToString().IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (!string.IsNullOrEmpty(cmd.ModuleCode) && cmd.ModuleCode.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) return true;
            if (cmd.Keywords.Any(k => k.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)) return true;
            return false;
        }

        private static int Score(SuiteCommand cmd, string query)
        {
            if (cmd.Title.Equals(query, StringComparison.OrdinalIgnoreCase)) return 100;
            if (cmd.Title.StartsWith(query, StringComparison.OrdinalIgnoreCase)) return 80;
            if (cmd.Title.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) return 60;
            if (cmd.Keywords.Any(k => k.Equals(query, StringComparison.OrdinalIgnoreCase))) return 50;
            if (cmd.Keywords.Any(k => k.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0)) return 40;
            if (cmd.Description.IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) return 30;
            if (cmd.Category.ToString().IndexOf(query, StringComparison.OrdinalIgnoreCase) >= 0) return 20;
            return 10;
        }
    }
}
