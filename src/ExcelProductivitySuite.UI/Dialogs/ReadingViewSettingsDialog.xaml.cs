using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using ExcelProductivitySuite.Core.Abstractions;
using ExcelProductivitySuite.Core.Models;
using ExcelProductivitySuite.Engine.Navigation;
using ExcelProductivitySuite.UI.Controls;

namespace ExcelProductivitySuite.UI.Dialogs
{
    public partial class ReadingViewSettingsDialog : BaseDialogWindow
    {
        private readonly IExcelContext? _context;
        private ReadingViewSettings _settings;
        private bool _isInitializing = true;

        public ReadingViewSettings Settings => _settings;

        public ReadingViewSettingsDialog() : this(null)
        {
        }

        public ReadingViewSettingsDialog(IExcelContext? context)
        {
            InitializeComponent();
            _context = context;
            _settings = context?.ReadingViewSettings?.Clone() ?? ReadingLayoutEngine.LoadSettings();

            Owner = Application.Current?.MainWindow;
            LoadSettingsToUi();
            _isInitializing = false;
            UpdatePreview();
        }

        private void LoadSettingsToUi()
        {
            switch (_settings.Shape)
            {
                case ReadingHighlightShape.RowOnly:
                    RbRowOnly.IsChecked = true;
                    break;
                case ReadingHighlightShape.ColumnOnly:
                    RbColOnly.IsChecked = true;
                    break;
                case ReadingHighlightShape.CellOnly:
                    RbCellOnly.IsChecked = true;
                    break;
                default:
                    RbCrosshair.IsChecked = true;
                    break;
            }

            if (_settings.Scope == ReadingHighlightScope.UsedRangeOnly)
            {
                RbScopeUsed.IsChecked = true;
            }
            else
            {
                RbScopeFull.IsChecked = true;
            }

            HexInput.Text = _settings.ColorHex;
            TransSlider.Value = _settings.Transparency * 100.0;
            BorderCheck.IsChecked = _settings.BorderEnabled;

            UpdateStatusBadge();
        }

        private void UpdateStatusBadge()
        {
            if (CurrentStatusBadge == null) return;
            bool isActive = _context?.IsReadingViewActive ?? _settings.IsEnabled;
            if (isActive)
            {
                CurrentStatusBadge.Text = "Status: Grid Focus is currently ON (Live)";
                CurrentStatusBadge.Foreground = new SolidColorBrush(Color.FromRgb(16, 124, 65)); // Emerald
            }
            else
            {
                CurrentStatusBadge.Text = "Status: Grid Focus is currently OFF";
                CurrentStatusBadge.Foreground = new SolidColorBrush(Color.FromRgb(100, 116, 139)); // Slate
            }
        }

        private void Shape_Changed(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;

            if (RbRowOnly?.IsChecked == true) _settings.Shape = ReadingHighlightShape.RowOnly;
            else if (RbColOnly?.IsChecked == true) _settings.Shape = ReadingHighlightShape.ColumnOnly;
            else if (RbCellOnly?.IsChecked == true) _settings.Shape = ReadingHighlightShape.CellOnly;
            else _settings.Shape = ReadingHighlightShape.RowAndColumn;

            UpdatePreview();
        }

        private void Scope_Changed(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;

            if (RbScopeUsed?.IsChecked == true) _settings.Scope = ReadingHighlightScope.UsedRangeOnly;
            else _settings.Scope = ReadingHighlightScope.EntireRowColumn;

            UpdatePreview();
        }

        private void ColorPreset_Click(object sender, RoutedEventArgs e)
        {
            if (sender is Button btn && btn.Tag is string hex)
            {
                HexInput.Text = hex;
            }
        }

        private void HexInput_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (_isInitializing) return;

            string hex = HexInput.Text?.Trim() ?? string.Empty;
            if (hex.Length >= 6)
            {
                try
                {
                    if (!hex.StartsWith("#")) hex = "#" + hex;
                    var color = (Color)ColorConverter.ConvertFromString(hex);
                    _settings.ColorHex = hex;
                    ColorIndicator.Background = new SolidColorBrush(color);
                    UpdatePreview();
                }
                catch { }
            }
        }

        private void TransSlider_ValueChanged(object sender, RoutedPropertyChangedEventArgs<double> e)
        {
            if (_isInitializing || TransValueText == null) return;

            int transPct = (int)TransSlider.Value;
            int opacityPct = 100 - transPct;
            TransValueText.Text = $"{transPct}% Transparent ({opacityPct}% Opacity)";
            _settings.Transparency = transPct / 100.0;

            UpdatePreview();
        }

        private void Border_Changed(object sender, RoutedEventArgs e)
        {
            if (_isInitializing) return;
            _settings.BorderEnabled = BorderCheck?.IsChecked == true;
            UpdatePreview();
        }

        private void UpdatePreview()
        {
            if (PreviewRowHighlight == null || PreviewColHighlight == null || PreviewCellBorder == null) return;

            try
            {
                var fillColor = (Color)ColorConverter.ConvertFromString(_settings.ColorHex);
                double opacity = Math.Max(0.05, Math.Min(1.0, 1.0 - _settings.Transparency));

                var brush = new SolidColorBrush(fillColor);
                PreviewRowHighlight.Background = brush;
                PreviewRowHighlight.Opacity = opacity;

                PreviewColHighlight.Background = brush;
                PreviewColHighlight.Opacity = opacity;

                // Shape visibility
                bool showRow = _settings.Shape == ReadingHighlightShape.RowAndColumn || _settings.Shape == ReadingHighlightShape.RowOnly;
                bool showCol = _settings.Shape == ReadingHighlightShape.RowAndColumn || _settings.Shape == ReadingHighlightShape.ColumnOnly;
                bool showCell = _settings.Shape == ReadingHighlightShape.CellOnly || _settings.Shape == ReadingHighlightShape.RowAndColumn;

                PreviewRowHighlight.Visibility = showRow ? Visibility.Visible : Visibility.Collapsed;
                PreviewColHighlight.Visibility = showCol ? Visibility.Visible : Visibility.Collapsed;

                if (_settings.BorderEnabled)
                {
                    PreviewCellBorder.Visibility = showCell ? Visibility.Visible : Visibility.Collapsed;
                    PreviewCellBorder.BorderBrush = new SolidColorBrush(fillColor);
                }
                else
                {
                    PreviewCellBorder.Visibility = Visibility.Collapsed;
                }
            }
            catch { }
        }

        private void Enable_Click(object sender, RoutedEventArgs e)
        {
            _settings.IsEnabled = true;
            ReadingLayoutEngine.SaveSettings(_settings);

            if (_context != null)
            {
                _context.ReadingViewSettings = _settings;
                _context.EnableReadingView(_settings);
            }

            DialogResult = true;
            Close();
        }

        private void Disable_Click(object sender, RoutedEventArgs e)
        {
            _settings.IsEnabled = false;
            ReadingLayoutEngine.SaveSettings(_settings);

            if (_context != null)
            {
                _context.DisableReadingView();
            }

            UpdateStatusBadge();
            _context?.ShowMessage("Grid Focus (Reading Layout) has been turned off.", "Grid Focus");
        }

        private void Close_Click(object sender, RoutedEventArgs e)
        {
            DialogResult = false;
            Close();
        }
    }
}
