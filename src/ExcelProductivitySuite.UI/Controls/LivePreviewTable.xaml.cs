using System.Collections;
using System.Windows;
using System.Windows.Controls;

namespace ExcelProductivitySuite.UI.Controls
{
    public partial class LivePreviewTable : UserControl
    {
        public static readonly DependencyProperty ItemsSourceProperty =
            DependencyProperty.Register(nameof(ItemsSource), typeof(IEnumerable), typeof(LivePreviewTable),
                new PropertyMetadata(null, OnItemsSourceChanged));

        public IEnumerable ItemsSource
        {
            get => (IEnumerable)GetValue(ItemsSourceProperty);
            set => SetValue(ItemsSourceProperty, value);
        }

        public LivePreviewTable()
        {
            InitializeComponent();
        }

        private static void OnItemsSourceChanged(DependencyObject d, DependencyPropertyChangedEventArgs e)
        {
            if (d is LivePreviewTable table)
            {
                table.PreviewGrid.ItemsSource = e.NewValue as IEnumerable;
            }
        }
    }
}
