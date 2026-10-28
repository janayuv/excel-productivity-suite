using System;
using System.Collections.Generic;
using ExcelProductivitySuite.Core.Abstractions;

namespace ExcelProductivitySuite.Engine.Ranges
{
    public class CommentFilterOptions
    {
        public string FilterText { get; set; } = "";
        public string FilterAuthor { get; set; } = "";
        public bool MatchCase { get; set; } = false;
    }

    /// <summary>
    /// Filters and inspects cell comments across worksheets.
    /// </summary>
    public static class CommentFilter
    {
        public static List<CommentInfo> Filter(IEnumerable<CommentInfo> comments, CommentFilterOptions options)
        {
            if (comments == null) throw new ArgumentNullException(nameof(comments));
            if (options == null) throw new ArgumentNullException(nameof(options));

            var results = new List<CommentInfo>();
            var comparison = options.MatchCase ? StringComparison.Ordinal : StringComparison.OrdinalIgnoreCase;

            foreach (var c in comments)
            {
                if (!string.IsNullOrEmpty(options.FilterAuthor))
                {
                    if (c.Author.IndexOf(options.FilterAuthor, comparison) < 0) continue;
                }

                if (!string.IsNullOrEmpty(options.FilterText))
                {
                    if (c.Text.IndexOf(options.FilterText, comparison) < 0) continue;
                }

                results.Add(c);
            }

            return results;
        }
    }
}
