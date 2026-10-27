using System.ComponentModel;

namespace ExcelProductivitySuite.UI.Models
{
    public class PreviewItem : INotifyPropertyChanged
    {
        private string _original = "";
        private string _result = "";

        public int RowIndex { get; set; }

        public string Original
        {
            get => _original;
            set { _original = value; OnPropertyChanged(nameof(Original)); }
        }

        public string Result
        {
            get => _result;
            set { _result = value; OnPropertyChanged(nameof(Result)); }
        }

        public event PropertyChangedEventHandler? PropertyChanged;
        protected virtual void OnPropertyChanged(string propertyName) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}
