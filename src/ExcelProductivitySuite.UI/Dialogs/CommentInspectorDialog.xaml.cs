using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Engine.Ranges;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class CommentInspectorDialog : BaseDialogWindow
    {
        private readonly IExcelContext _context;
        private List<CommentInfo> _allComments = new();
        private List<CommentInfo> _filteredComments = new();

        public CommentInspectorDialog(IExcelContext context)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            InitializeComponent();

            LoadComments();
        }

        private void LoadComments()
        {
            try
            {
                var sheets = _context.GetWorksheetNames();
                _allComments.Clear();

                foreach (var sheet in sheets)
                {
                    var comments = _context.GetWorksheetComments(sheet);
                    if (comments != null)
                    {
                        _allComments.AddRange(comments);
                    }
                }

                ApplyFilter();
            }
            catch (Exception ex)
            {
                StatusText.Text = "Error loading comments: " + ex.Message;
            }
        }

        private void ApplyFilter()
        {
            var options = new CommentFilterOptions
            {
                FilterText = FilterTextBox?.Text?.Trim() ?? "",
                FilterAuthor = AuthorBox?.Text?.Trim() ?? "",
                MatchCase = false
            };

            _filteredComments = CommentFilter.Filter(_allComments, options);

            CommentsGrid.ItemsSource = null;
            CommentsGrid.ItemsSource = _filteredComments;

            StatusText.Text = $"Displaying {_filteredComments.Count} comment(s) (Total: {_allComments.Count} in workbook).";
        }

        private void FilterChanged(object sender, TextChangedEventArgs e)
        {
            ApplyFilter();
        }

        private void NavigateButton_Click(object sender, RoutedEventArgs e)
        {
            if (CommentsGrid.SelectedItem is CommentInfo selected)
            {
                if (selected.Row > 0 && selected.Column > 0)
                {
                    _context.SelectRange(selected.Row, selected.Column, 1, 1);
                    Close();
                }
            }
            else
            {
                _context.ShowMessage("Please select a comment from the list first.", "Comment Inspector");
            }
        }

        private void CloseButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
